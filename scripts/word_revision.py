#!/usr/bin/env python3
"""Deterministic DOCX text revisions and comments. Python 3.10+, lxml."""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from copy import deepcopy
from datetime import datetime, timezone
import difflib
import hashlib
import io
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import zipfile

from lxml import etree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
XML = 'http://www.w3.org/XML/1998/namespace'
NS = {'w': W}
MAIN = 'word/document.xml'
SETTINGS = 'word/settings.xml'
COMMENTS = 'word/comments.xml'
RELS = 'word/_rels/document.xml.rels'
TYPES = '[Content_Types].xml'
REV_TAGS = {'ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange',
            'tblPrChange', 'trPrChange', 'tcPrChange', 'sectPrChange',
            'tblGridChange', 'numberingChange', 'cellIns', 'cellDel', 'cellMerge'}
VERSION = '0.1.0'


class Conflict(Exception):
    def __init__(self, message, details=None):
        super().__init__(message)
        self.details = details or []


def q(name):
    return '{' + W + '}' + name


def sha(data):
    return hashlib.sha256(data).hexdigest()


def xml(data):
    parser = ET.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    root = ET.fromstring(data, parser)
    if root.getroottree().docinfo.doctype:
        raise Conflict('XML DTD requires a separate document conversion')
    return root


def serialized(root):
    return ET.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def canon(root):
    """Prefix-independent XML identity, including attributes and content."""
    if root is None:
        return None
    return (root.tag, tuple(sorted(root.attrib.items())), root.text or '',
            tuple((canon(c), c.tail or '') for c in root))


def simple_run(run):
    return (run.tag == q('r') and len(run.findall(q('t'))) > 0
            and all(c.tag in (q('rPr'), q('t')) for c in run)
            and all(set(t.attrib) <= {'{' + XML + '}space'} for t in run.findall(q('t'))))


def run_signature(run):
    return (tuple(sorted(run.attrib.items())), canon(run.find(q('rPr'))))


def run_text(run):
    return ''.join(t.text or '' for t in run.findall(q('t')))


def normalized(root):
    """Semantic identity permitting only splitting/joining equivalent text runs."""
    def walk(node):
        children = []
        for child in node:
            if simple_run(child):
                item = ('text-run', run_signature(child), run_text(child))
                if children and children[-1][0] == 'text-run' and children[-1][1] == item[1]:
                    previous = children.pop()
                    item = ('text-run', item[1], previous[2] + item[2])
                children.append(item)
            else:
                children.append(('node', walk(child), child.tail or ''))
        return (node.tag, tuple(sorted(node.attrib.items())), node.text or '', tuple(children))
    return walk(root)


class Package:
    def __init__(self, data):
        self.data = data
        self.digest = sha(data)
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                self.infos = z.infolist()
                names = [i.filename for i in self.infos]
                if len(names) != len(set(names)):
                    raise Conflict('Duplicate ZIP part names')
                if sum(i.file_size for i in self.infos) > 256 * 1024 * 1024:
                    raise Conflict('Document exceeds the 256 MiB unpacked limit')
                self.parts = {n: z.read(n) for n in names}
                self.zip_comment = z.comment
            for n in (MAIN, TYPES):
                if n not in self.parts:
                    raise Conflict('Expected DOCX part missing: ' + n)
            self.roots = {n: xml(b) for n, b in self.parts.items()
                          if n.endswith(('.xml', '.rels'))}
            self.root = self.roots[MAIN]
            if self.root.tag != q('document'):
                raise Conflict('Convert this document to transitional .docx before editing')
            self.paragraphs = self.root.findall('.//' + q('p'))
        except (zipfile.BadZipFile, ET.XMLSyntaxError, KeyError) as exc:
            raise Conflict('Unreadable DOCX: ' + str(exc)) from exc

    @classmethod
    def read(cls, path):
        return cls(Path(path).read_bytes())

    def pack(self, overrides):
        result = io.BytesIO()
        with zipfile.ZipFile(result, 'w', zipfile.ZIP_DEFLATED) as z:
            z.comment = self.zip_comment
            for info in self.infos:
                z.writestr(info, overrides.get(info.filename, self.parts[info.filename]))
            for name, content in overrides.items():
                if name not in self.parts:
                    z.writestr(name, content)
        return result.getvalue()


def text_view(node, view='current'):
    """Text projection; paragraph/row formatting revisions remain separate metadata."""
    excluded = {'del', 'moveFrom'} if view == 'current' else {'ins', 'moveTo'}
    def walk(e):
        if e.tag in {q(n) for n in excluded}:
            return ''
        if e is not node and e.tag == q('p'):
            return ''  # Nested textboxes have their own paragraph records.
        if e.tag in (q('t'), q('delText')):
            return e.text or ''
        if e.tag == q('tab'):
            return '\t'
        if e.tag in (q('br'), q('cr')):
            return '\n'
        return ''.join(walk(c) for c in e)
    return walk(node)


def pid(index):
    return f'p{index + 1:06d}'


def existing_revisions(pkg):
    records = []
    for part, root in pkg.roots.items():
        if not part.startswith('word/'):
            continue
        for e in root.iter():
            if isinstance(e.tag, str) and e.tag.startswith('{' + W + '}'):
                kind = ET.QName(e).localname
                if kind in REV_TAGS:
                    records.append({'part': part, 'kind': kind, 'id': e.get(q('id')),
                                    'author': e.get(q('author')), 'date': e.get(q('date')),
                                    'text': text_view(e, 'current') or text_view(e, 'original')})
    return records


def inspect_document(pkg, query=None, paragraph=None, summary=False):
    blocks = []
    for i, p in enumerate(pkg.paragraphs):
        current, original = text_view(p), text_view(p, 'original')
        if paragraph and pid(i) != paragraph:
            continue
        if query and query.casefold() not in (current + '\n' + original).casefold():
            continue
        blocks.append({'id': pid(i), 'text': current, 'text_sha256': sha(current.encode()),
                       'original_text': original,
                       'in_table': any(a.tag == q('tc') for a in p.iterancestors()),
                       'comment_ids': p.xpath('.//w:commentReference/@w:id', namespaces=NS),
                       'style': p.xpath('string(w:pPr/w:pStyle/@w:val)', namespaces=NS)})
    comments = []
    if COMMENTS in pkg.roots:
        for c in pkg.roots[COMMENTS]:
            comments.append({'id': c.get(q('id')), 'author': c.get(q('author')),
                             'date': c.get(q('date')),
                             'text': '\n'.join(text_view(p) for p in c.findall('.//' + q('p')))})
    others = {}
    for part, root in pkg.roots.items():
        if re.fullmatch(r'word/(header\d+|footer\d+|footnotes|endnotes)\.xml', part):
            others[part] = [text_view(p) for p in root.findall('.//' + q('p'))]
    result = {'source_sha256': pkg.digest, 'version': VERSION,
              'paragraph_count': len(pkg.paragraphs), 'comments': comments,
              'revisions': existing_revisions(pkg), 'other_stories': others,
              'projection': 'text insertions/deletions; structural revisions listed separately',
              'fields': pkg.root.xpath('//w:instrText/text() | //w:fldSimple/@w:instr', namespaces=NS)}
    if not summary:
        result['paragraphs'] = blocks
    return result


def validate_plan(pkg, plan):
    errors = []
    if plan.get('source_sha256') != pkg.digest:
        errors.append({'code': 'source_changed', 'message': 'Inspect the current source and rebuild the plan'})
    if any(n.startswith('_xmlsignatures/') for n in pkg.parts):
        errors.append({'code': 'signed_package', 'message': 'Use an unsigned editing copy'})
    settings = pkg.roots.get(SETTINGS)
    if settings is not None:
        protection = settings.find(q('documentProtection'))
        if protection is not None and protection.get(q('enforcement'), 'false') not in ('0', 'false', 'off'):
            errors.append({'code': 'document_protection', 'message': 'Use an editable document copy'})
    operations = plan.get('operations')
    if not isinstance(operations, list) or not operations:
        errors.append({'code': 'empty_plan', 'message': 'Provide at least one operation'})
        return [], errors
    resolved, occupied = [], defaultdict(list)
    for n, op in enumerate(operations):
        try:
            if not isinstance(op, dict):
                raise Conflict('Operation must be a JSON object')
            kind = op.get('type', 'replace')
            if kind not in ('replace', 'comment'):
                raise Conflict('Use type replace or comment')
            allowed_keys = ({'type', 'paragraph', 'before', 'after', 'author', 'comment'}
                            if kind == 'replace' else {'type', 'paragraph', 'quote', 'text', 'author'})
            unknown = set(op) - allowed_keys
            if unknown:
                raise Conflict('Unknown operation fields: ' + ', '.join(sorted(unknown)))
            author = op.get('author', plan.get('author'))
            if not isinstance(author, str) or not author.strip():
                raise Conflict('Provide the review author')
            name = op.get('paragraph', '')
            if not isinstance(name, str) or not re.fullmatch(r'p\d{6}', name):
                raise Conflict('Use a paragraph id from inspect')
            index = int(name[1:]) - 1
            if not 0 <= index < len(pkg.paragraphs):
                raise Conflict('Paragraph id is outside the document')
            p = pkg.paragraphs[index]
            old = op.get('before') if kind == 'replace' else op.get('quote')
            new = op.get('after') if kind == 'replace' else op.get('text')
            if not isinstance(old, str) or not old:
                raise Conflict('Provide a nonempty before/quote anchor')
            if not isinstance(new, str) or (kind == 'comment' and not new.strip()):
                raise Conflict('Provide after/text as a string')
            # These operations affect inline text. Paragraph structure stays intact.
            if kind == 'replace' and any(c in new + old for c in '\r\n\t'):
                raise Conflict('Use an inline text anchor without tabs or paragraph breaks')
            if kind == 'replace' and old == new:
                raise Conflict('Replacement has identical before and after text')
            annotation = op.get('comment')
            if annotation is not None and (not isinstance(annotation, str) or not annotation.strip()):
                raise Conflict('Replacement comment must be a nonempty string')
            for value in (old, new, author, annotation or ''):
                ET.Element(q('t')).text = value  # XML character validation.
            content = text_view(p)
            hits = [m.start() for m in re.finditer('(?=' + re.escape(old) + ')', content)]
            if len(hits) != 1:
                raise Conflict(f'Anchor matches {len(hits)} times; extend it to unique context')
            start, end = hits[0], hits[0] + len(old)
            if any(start < b and end > a for a, b in occupied[index]):
                raise Conflict('Operations overlap in the source paragraph')
            eligible(p, start, end)
            occupied[index].append((start, end))
            resolved.append({'number': n + 1, 'type': kind, 'paragraph': name, 'index': index,
                             'start': start, 'end': end, 'before': old, 'after': new,
                             'author': author, 'comment': annotation})
        except (Conflict, ValueError, TypeError) as exc:
            errors.append({'operation': n + 1, 'code': 'target_conflict', 'message': str(exc)})
    return resolved, errors


def spans(p):
    result, offset = [], 0
    for child in p:
        text = text_view(child)
        result.append((child, offset, offset + len(text)))
        offset += len(text)
    return result


def eligible(p, start, end):
    # Property/paragraph/row revisions require Word's richer revision semantics.
    active_moves = set()
    for element in p.getroottree().getroot().iter():
        if element is p:
            break
        if element.tag in (q('moveFromRangeStart'), q('moveToRangeStart')):
            active_moves.add((element.tag, element.get(q('id'))))
        elif element.tag in (q('moveFromRangeEnd'), q('moveToRangeEnd')):
            start_tag = element.tag.replace('RangeEnd', 'RangeStart')
            active_moves.discard((start_tag, element.get(q('id'))))
    if active_moves:
        raise Conflict('Paragraph belongs to an existing move range')
    for a in [p, *p.iterancestors()]:
        if a.tag in {q(n) for n in ('ins', 'del', 'moveFrom', 'moveTo', 'sdt', 'txbxContent')}:
            raise Conflict('Anchor belongs to a revision, content control, or text box')
        props = {'p': 'pPr', 'tr': 'trPr', 'tc': 'tcPr'}.get(ET.QName(a).localname)
        if props:
            pr = a.find(q(props))
            if pr is not None and any(e.tag in {q(x) for x in REV_TAGS} for e in pr.iter()):
                raise Conflict('Anchor has a structural or formatting revision')
    if p.xpath('.//w:fldChar | .//w:fldSimple | .//w:instrText | .//w:moveFromRangeStart | .//w:moveToRangeStart', namespaces=NS):
        raise Conflict('Paragraph contains a field or move range; use a native Word edit for this target')
    covered = [(r, a, b) for r, a, b in spans(p) if a < end and b > start]
    if not covered or not all(simple_run(r) for r, _, _ in covered):
        raise Conflict('Anchor crosses an existing revision or a structured text object')
    for child, a, b in spans(p):
        if start < a < end and a == b and child.tag != q('pPr'):
            raise Conflict('Anchor crosses a bookmark, comment, or another XML boundary')
    if any(r.find('.//' + q('rPrChange')) is not None for r, _, _ in covered):
        raise Conflict('Anchor contains an earlier formatting revision')


def fragment(run, text, deleted=False):
    result = deepcopy(run)
    for child in list(result):
        if child.tag != q('rPr'):
            result.remove(child)
    t = ET.SubElement(result, q('delText') if deleted else q('t'))
    t.set('{' + XML + '}space', 'preserve')
    t.text = text
    result.tail = None
    return result


class Ids:
    def __init__(self, pkg):
        numbers = [int(v) for r in pkg.roots.values() for v in r.xpath('//@w:id', namespaces=NS)
                   if re.fullmatch(r'\d+', v)]
        self.value = max(numbers, default=-1) + 1

    def next(self):
        result = str(self.value)
        self.value += 1
        return result


def revision(kind, runs, author, date, ids, new_ids):
    e = ET.Element(q(kind))
    ident = ids.next()
    e.attrib.update({q('id'): ident, q('author'): author, q('date'): date})
    e.extend(runs)
    new_ids.append(ident)
    return e


def change_span(p, start, end, replacement, op, date, ids, new_ids):
    entries = spans(p)
    if start == end:
        # Prefer the left run at a boundary; both are ordinary editable runs.
        entry = next(((r, a, b) for r, a, b in entries if simple_run(r) and a < start <= b), None)
        if entry is None:
            entry = next(((r, a, b) for r, a, b in entries if simple_run(r) and a <= start < b), None)
        if entry is None:
            raise Conflict('Insertion anchor requires a text run')
        r, a, b = entry
        text = run_text(r)
        nodes = []
        if start > a:
            nodes.append(fragment(r, text[:start-a]))
        nodes.append(revision('ins', [fragment(r, replacement)], op['author'], date, ids, new_ids))
        if start < b:
            nodes.append(fragment(r, text[start-a:]))
        at = p.index(r)
        p.remove(r)
        for node in reversed(nodes):
            p.insert(at, node)
        return
    covered = [(r, a, b) for r, a, b in entries if a < end and b > start]
    if not all(simple_run(r) for r, _, _ in covered):
        raise Conflict('Internal target changed during patch construction')
    first, left, _ = covered[0]
    last, last_start, right = covered[-1]
    nodes = []
    if start > left:
        nodes.append(fragment(first, run_text(first)[:start-left]))
    deleted = [fragment(r, run_text(r)[max(start,a)-a:min(end,b)-a], True) for r,a,b in covered]
    nodes.append(revision('del', deleted, op['author'], date, ids, new_ids))
    if replacement:
        nodes.append(revision('ins', [fragment(first, replacement)], op['author'], date, ids, new_ids))
    if end < right:
        nodes.append(fragment(last, run_text(last)[end-last_start:]))
    at = p.index(first)
    for r, _, _ in covered:
        p.remove(r)
    for node in reversed(nodes):
        p.insert(at, node)


def split_boundary(p, offset):
    """Return child index at an exact visible-text boundary, preserving run format."""
    for r, start, end in spans(p):
        if start == offset and end > start:
            return p.index(r)
        if start < offset < end:
            if not simple_run(r):
                raise Conflict('Comment boundary requires an ordinary text run')
            at = p.index(r)
            text = run_text(r)
            p.remove(r)
            p.insert(at, fragment(r, text[:offset-start]))
            p.insert(at+1, fragment(r, text[offset-start:]))
            return at+1
        if end == offset and end > start:
            return p.index(r)+1
    raise Conflict('Comment anchor is outside paragraph text')


def add_comment(p, op, comments, date, ids, comment_ids):
    ident = ids.next()
    end_at = split_boundary(p, op['end'])
    end = ET.Element(q('commentRangeEnd'), {q('id'): ident})
    p.insert(end_at, end)
    ref = ET.Element(q('r'))
    ET.SubElement(ref, q('commentReference'), {q('id'): ident})
    p.insert(end_at+1, ref)
    start_at = split_boundary(p, op['start'])
    p.insert(start_at, ET.Element(q('commentRangeStart'), {q('id'): ident}))
    c = ET.SubElement(comments, q('comment'), {q('id'): ident, q('author'): op['author'], q('date'): date})
    for line in op['after'].splitlines() or ['']:
        para = ET.SubElement(c, q('p'))
        run = ET.SubElement(para, q('r'))
        t = ET.SubElement(run, q('t'), {'{' + XML + '}space': 'preserve'})
        t.text = line
    comment_ids.append(ident)


def wire_part(rels, types, kind):
    target = kind + '.xml'
    matches = [e for e in rels if e.get('Type') == REL + kind]
    if matches:
        if len(matches) != 1 or matches[0].get('Target') not in (target, '/word/' + target):
            raise Conflict('Nonstandard ' + kind + ' part relationship requires conversion')
    else:
        used = {e.get('Id') for e in rels}
        i = 1
        while 'rId' + str(i) in used:
            i += 1
        ET.SubElement(rels, '{' + R + '}Relationship', Id='rId' + str(i), Type=REL + kind, Target=target)
    part = '/word/' + target
    mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml'
    matches = [e for e in types if e.get('PartName') == part]
    if matches:
        if len(matches) != 1 or matches[0].get('ContentType') != mime:
            raise Conflict('Unexpected content type for ' + part)
    else:
        ET.SubElement(types, '{' + CT + '}Override', PartName=part, ContentType=mime)


def enable_tracking(settings):
    track = settings.find(q('trackRevisions'))
    if track is None:
        track = ET.Element(q('trackRevisions'))
        # CT_Settings sequence: trackRevisions follows revisionView.
        before = {'writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime',
                  'doNotDisplayPageBoundaries', 'displayBackgroundShape', 'printPostScriptOverText',
                  'printFractionalCharacterWidth', 'printFormsData', 'embedTrueTypeFonts',
                  'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins',
                  'alignBordersAndEdges', 'bordersDoNotSurroundHeader', 'bordersDoNotSurroundFooter',
                  'gutterAtTop', 'hideSpellingErrors', 'hideGrammaticalErrors', 'activeWritingStyle',
                  'proofState', 'formsDesign', 'attachedTemplate', 'linkStyles', 'stylePaneFormatFilter',
                  'stylePaneSortMethod', 'documentType', 'mailMerge', 'revisionView'}
        index = 0
        for i, child in enumerate(settings):
            if ET.QName(child).localname in before:
                index = i + 1
        settings.insert(index, track)
    track.set(q('val'), 'true')


def undo_batch(root, revision_ids, comment_ids, accept=False):
    result = deepcopy(root)
    revisions, comments = set(revision_ids), set(comment_ids)
    for e in list(result.iter()):
        ident = e.get(q('id'))
        parent = e.getparent()
        if parent is None:
            continue
        if e.tag in (q('commentRangeStart'), q('commentRangeEnd')) and ident in comments:
            parent.remove(e)
        elif e.tag == q('commentReference') and ident in comments:
            parent.remove(e)
            if len(parent) == 0:
                parent.getparent().remove(parent)
        elif e.tag in (q('ins'), q('del')) and ident in revisions:
            keep = (e.tag == q('ins')) == accept
            at = parent.index(e)
            if keep:
                for child in list(e):
                    if not accept:
                        for t in child.iter(q('delText')):
                            t.tag = q('t')
                    parent.insert(at, child)
                    at += 1
            parent.remove(e)
    return result


def build(pkg, plan, date=None):
    ops, errors = validate_plan(pkg, plan)
    if errors:
        raise Conflict('Preflight found conflicts; no document written', errors)
    date = date or datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    root = deepcopy(pkg.root)
    paras = root.findall('.//' + q('p'))
    comments = deepcopy(pkg.roots[COMMENTS]) if COMMENTS in pkg.roots else ET.Element(q('comments'), nsmap={'w': W})
    ids = Ids(pkg)
    revision_ids, comment_ids = [], []
    expected = [text_view(p) for p in pkg.paragraphs]
    for op in sorted(ops, key=lambda x: (x['index'], x['start']), reverse=True):
        p = paras[op['index']]
        if op['type'] == 'comment':
            add_comment(p, op, comments, date, ids, comment_ids)
        else:
            if op.get('comment'):
                annotation = dict(op, after=op['comment'])
                add_comment(p, annotation, comments, date, ids, comment_ids)
            text = expected[op['index']]
            expected[op['index']] = text[:op['start']] + op['after'] + text[op['end']:]
            matcher = difflib.SequenceMatcher(a=op['before'], b=op['after'], autojunk=False)
            for tag, i, j, a, b in reversed(matcher.get_opcodes()):
                if tag != 'equal':
                    change_span(p, op['start']+i, op['start']+j, op['after'][a:b], op, date, ids, revision_ids)
    overrides = {MAIN: serialized(root)}
    rels = deepcopy(pkg.roots[RELS]) if RELS in pkg.roots else ET.Element('{' + R + '}Relationships', nsmap={None: R})
    types = deepcopy(pkg.roots[TYPES])
    if revision_ids:
        settings = deepcopy(pkg.roots[SETTINGS]) if SETTINGS in pkg.roots else ET.Element(q('settings'), nsmap={'w': W})
        enable_tracking(settings)
        wire_part(rels, types, 'settings')
        overrides[SETTINGS] = serialized(settings)
    if comment_ids:
        wire_part(rels, types, 'comments')
        overrides[COMMENTS] = serialized(comments)
    if canon(rels) != canon(pkg.roots.get(RELS)):
        overrides[RELS] = serialized(rels)
    if canon(types) != canon(pkg.roots[TYPES]):
        overrides[TYPES] = serialized(types)
    output = pkg.pack(overrides)
    report = audit(pkg, Package(output), plan, revision_ids, comment_ids, expected)
    report.update({'date': date, 'operations': [{k:v for k,v in op.items() if k != 'index'} for op in ops],
                   'revision_ids': revision_ids, 'comment_ids': comment_ids,
                   'format_inheritance': 'new text uses the first affected run; retained text keeps its own properties'})
    return output, report


def audit(source, output, plan, revision_ids, comment_ids, expected=None):
    ops, errors = validate_plan(source, plan)
    if errors:
        raise Conflict('Source or plan no longer matches', errors)
    if expected is None:
        expected = [text_view(p) for p in source.paragraphs]
        for op in sorted(ops, key=lambda x: (x['index'], x['start']), reverse=True):
            if op['type'] == 'replace':
                t = expected[op['index']]
                expected[op['index']] = t[:op['start']] + op['after'] + t[op['end']:]
    changed = sorted(n for n in set(source.parts) | set(output.parts)
                     if source.parts.get(n) != output.parts.get(n))
    allowed = {MAIN, SETTINGS, COMMENTS, RELS, TYPES}
    checks = {
        'planned_current_text': [text_view(p) for p in output.paragraphs] == expected,
        'reject_batch_restores_formatted_content': normalized(undo_batch(output.root, revision_ids, comment_ids)) == normalized(source.root),
        'unchanged_other_parts': not (set(changed) - allowed),
        'paragraph_count': len(source.paragraphs) == len(output.paragraphs),
    }
    # Existing revision elements stay byte-semantically identical, including nesting.
    def revisions(root, excluded=()):
        return Counter(repr(canon(e)) for e in root.iter()
                       if e.tag in {q(x) for x in REV_TAGS} and e.get(q('id')) not in excluded)
    checks['existing_revisions'] = revisions(source.root) == revisions(output.root, set(revision_ids))
    old_comments = source.roots.get(COMMENTS)
    new_comments = output.roots.get(COMMENTS)
    preserved = deepcopy(new_comments) if new_comments is not None else None
    if preserved is not None:
        for e in list(preserved):
            if e.get(q('id')) in set(comment_ids):
                preserved.remove(e)
    checks['existing_comments'] = (canon(preserved) == canon(old_comments) if old_comments is not None
                                    else preserved is None or len(preserved) == 0)
    for ident in comment_ids:
        checks['comment_anchor_' + ident] = all(len(output.root.xpath('.//w:' + tag + '[@w:id=$id]', namespaces=NS, id=ident)) == 1
                                                for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference'))
    new_revs = [e for e in output.root.iter() if e.tag in (q('ins'), q('del')) and e.get(q('id')) in set(revision_ids)]
    checks['new_revision_ids'] = len(new_revs) == len(set(revision_ids)) == len(revision_ids)
    checks['new_revision_authors'] = all(e.get(q('author')) in {o['author'] for o in ops} for e in new_revs)
    if revision_ids:
        setting = output.roots.get(SETTINGS)
        tracking = setting.find(q('trackRevisions')) if setting is not None else None
        checks['tracking_enabled'] = tracking is not None and tracking.get(q('val'), 'true') in ('true','1','on')
    report = {'ok': all(checks.values()), 'source_sha256': source.digest, 'output_sha256': output.digest,
              'checks': checks, 'changed_parts': changed, 'operation_count': len(ops),
              'verification': {'content_and_structure': 'passed' if all(checks.values()) else 'failed',
                               'layout': 'not_requested', 'native_word': 'not_run', 'full_schema': 'not_run'}}
    if not report['ok']:
        raise Conflict('Verification failed; no document written', [report])
    return report


def verify(source, output, plan, report):
    if not isinstance(report, dict):
        raise Conflict('Receipt must be a JSON object')
    for key in ('revision_ids', 'comment_ids'):
        if not isinstance(report.get(key), list) or not all(isinstance(x, str) for x in report[key]):
            raise Conflict('Receipt requires a string list: ' + key)
    if not isinstance(report.get('date'), str):
        raise Conflict('Receipt requires its recorded date')
    if report.get('source_sha256') != source.digest or report.get('output_sha256') != output.digest:
        raise Conflict('Verification receipt does not match these files')
    actual = audit(source, output, plan, report['revision_ids'], report['comment_ids'])
    # Rebuild independently from the supplied plan and recorded timestamp to cover
    # metadata, comments, relationships and settings, beyond the text projections.
    expected_data, _ = build(source, plan, report['date'])
    expected = Package(expected_data)
    if expected.parts != output.parts:
        raise Conflict('Output differs from the recorded plan')
    actual['checks']['replayed_plan_matches_all_parts'] = True
    return actual


def publish_new(path, data):
    """Publish a completed file without replacing an existing path."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix='.word-revision-', dir=path.parent)
    try:
        with os.fdopen(fd, 'wb') as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.link(tmp, path)  # Atomic and fails if the destination already exists.
    finally:
        os.unlink(tmp)


def apply_file(source_path, plan, output_path, receipt_path):
    source_path, output_path, receipt_path = map(Path, (source_path, output_path, receipt_path))
    if len({p.resolve() for p in (source_path, output_path, receipt_path)}) != 3:
        raise Conflict('Use separate source, revised document, and receipt paths')
    if output_path.exists() or receipt_path.exists():
        raise Conflict('Choose fresh output and receipt paths')
    source = Package.read(source_path)
    output, report = build(source, plan)
    if sha(source_path.read_bytes()) != source.digest:
        raise Conflict('Source changed during processing; inspect its current version')
    # A receipt without its document is recoverable; it is removed if publication fails.
    publish_new(receipt_path, (json.dumps(report, ensure_ascii=False, indent=2)+'\n').encode())
    try:
        publish_new(output_path, output)
    except BaseException:
        receipt_path.unlink()
        raise
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--version', action='version', version=VERSION)
    commands = parser.add_subparsers(dest='command', required=True)
    inspect = commands.add_parser('inspect', help='Extract text, revisions, comments and source fingerprint')
    inspect.add_argument('source')
    inspect.add_argument('--query')
    inspect.add_argument('--paragraph')
    inspect.add_argument('--summary', action='store_true')
    for name in ('check', 'apply', 'verify'):
        sub = commands.add_parser(name)
        sub.add_argument('source')
        sub.add_argument('--plan', required=True)
        if name == 'apply':
            sub.add_argument('--out', required=True)
            sub.add_argument('--receipt', required=True)
        if name == 'verify':
            sub.add_argument('output')
            sub.add_argument('--receipt', required=True)
    args = parser.parse_args()
    try:
        if args.command == 'inspect':
            result = inspect_document(Package.read(args.source), args.query, args.paragraph, args.summary)
        else:
            plan = json.loads(Path(args.plan).read_text(encoding='utf-8'))
            if not isinstance(plan, dict):
                raise Conflict('Plan must be a JSON object')
            if args.command == 'check':
                # Build and audit in memory, including package relationship checks.
                _, report = build(Package.read(args.source), plan)
                result = {'ok': True, 'operation_count': report['operation_count'], 'checks': report['checks'], 'writes': 0}
            elif args.command == 'apply':
                result = apply_file(args.source, plan, args.out, args.receipt)
            else:
                receipt = json.loads(Path(args.receipt).read_text(encoding='utf-8'))
                result = verify(Package.read(args.source), Package.read(args.output), plan, receipt)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except (Conflict, OSError, ValueError, KeyError, TypeError) as exc:
        print(json.dumps({'ok': False, 'message': str(exc), 'conflicts': getattr(exc, 'details', [])},
                         ensure_ascii=False, indent=2))
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
