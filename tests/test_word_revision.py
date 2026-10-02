"""Synthetic fixtures exercise public behavior; no personal documents are used."""
import importlib.util
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

SCRIPT = Path(__file__).resolve().parents[1] / 'scripts' / 'word_revision.py'
spec = importlib.util.spec_from_file_location('word_revision', SCRIPT)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


def package(body, extras=None, settings=None):
    parts = {
        '[Content_Types].xml': f'''<Types xmlns="{m.CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'''.encode(),
        '_rels/.rels': f'<Relationships xmlns="{m.R}"><Relationship Id="rId1" Type="{m.REL}officeDocument" Target="word/document.xml"/></Relationships>'.encode(),
        m.MAIN: f'<w:document xmlns:w="{m.W}"><w:body>{body}<w:sectPr/></w:body></w:document>'.encode(),
        'word/media/image1.bin': b'unchanged synthetic media',
    }
    if settings is not None:
        parts[m.SETTINGS] = settings.encode()
    parts.update(extras or {})
    b = io.BytesIO()
    with zipfile.ZipFile(b, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, value in parts.items():
            z.writestr(name, value)
    return m.Package(b.getvalue())


def run(text, props=''):
    return f'<w:r>{"<w:rPr>" + props + "</w:rPr>" if props else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'


def plan(pkg, *ops):
    return {'source_sha256': pkg.digest, 'author': 'Reviewer', 'operations': list(ops)}


def replace(before, after, paragraph='p000001', **kw):
    return {'paragraph': paragraph, 'before': before, 'after': after, **kw}


class Revisions(unittest.TestCase):
    def build(self, pkg, p):
        out, receipt = m.build(pkg, p, '2026-01-01T00:00:00Z')
        self.assertTrue(receipt['ok'])
        output = m.Package(out)
        self.assertTrue(m.verify(pkg, output, p, receipt)['ok'])
        return output, receipt

    def test_cross_run_minimal_revision_preserves_formats(self):
        p = package('<w:p>' + run('重复', '<w:b/>') + run('30', '<w:i/>') + run('次评分。') + '</w:p>')
        out, r = self.build(p, plan(p, replace('重复30次评分', '重复10次评分')))
        self.assertEqual(m.text_view(out.paragraphs[0]), '重复10次评分。')
        self.assertEqual(out.root.xpath('//w:delText/text()', namespaces=m.NS), ['3'])
        self.assertEqual(out.root.xpath('//w:ins/w:r/w:t/text()', namespaces=m.NS), ['1'])
        self.assertEqual(len(out.root.xpath('//w:ins/w:r/w:rPr/w:i', namespaces=m.NS)), 1)
        self.assertEqual(out.parts['word/media/image1.bin'], p.parts['word/media/image1.bin'])

    def test_cross_run_full_replacement(self):
        p = package('<w:p>' + run('alpha', '<w:b/>') + run('beta', '<w:i/>') + '</w:p>')
        out, r = self.build(p, plan(p, replace('alphabeta', '中文')))
        self.assertEqual(m.text_view(out.paragraphs[0]), '中文')
        deleted = out.root.xpath('//w:del/w:r/w:rPr/*', namespaces=m.NS)
        self.assertEqual([e.tag for e in deleted], [m.q('b'), m.q('i')])

    def test_insert_and_delete(self):
        for old, new in [('abc', 'Xabc'), ('abc', 'abcX'), ('abc', 'aXbc'), ('abc', '')]:
            with self.subTest(new=new):
                p = package('<w:p>' + run(old) + '</w:p>')
                out, _ = self.build(p, plan(p, replace(old, new)))
                self.assertEqual(m.text_view(out.paragraphs[0]), new)

    def test_multiple_edits_one_run_use_source_offsets(self):
        p = package('<w:p>' + run('one two three') + '</w:p>')
        out, _ = self.build(p, plan(p, replace('one', '1'), replace('three', '333333')))
        self.assertEqual(m.text_view(out.paragraphs[0]), '1 two 333333')

    def test_old_revisions_preserved(self):
        old = '<w:ins w:id="5000" w:author="Teacher" w:date="2025-01-01T00:00:00Z">' + run('旧增补') + '</w:ins>'
        p = package('<w:p>' + old + run(' 修改30次') + '</w:p>')
        out, receipt = self.build(p, plan(p, replace('30', '10')))
        self.assertEqual(m.canon(p.root.find('.//' + m.q('ins'))), m.canon(out.root.find('.//' + m.q('ins'))))
        self.assertTrue(all(int(i) > 5000 for i in receipt['revision_ids']))
        with self.assertRaises(m.Conflict):
            m.build(p, plan(p, replace('旧增补', '新内容')))

    def test_table_and_another_paragraph(self):
        p = package('<w:p>'+run('unchanged')+'</w:p><w:tbl><w:tr><w:tc><w:p>'+run('30')+'</w:p></w:tc></w:tr></w:tbl>')
        out, _ = self.build(p, plan(p, replace('30', '10', 'p000002')))
        self.assertEqual(m.text_view(out.paragraphs[1]), '10')
        self.assertTrue(m.inspect_document(out)['paragraphs'][1]['in_table'])

    def test_new_comment_wires_package(self):
        p = package('<w:p>' + run('one two', '<w:b/>') + run(' three') + '</w:p>')
        op = {'type':'comment','paragraph':'p000001','quote':'two three','text':'核对样本\nCheck sample'}
        out, r = self.build(p, plan(p, op))
        self.assertEqual(m.text_view(out.paragraphs[0]), 'one two three')
        self.assertIn(m.COMMENTS, out.parts)
        self.assertIn(b'comments.xml', out.parts[m.RELS])
        self.assertEqual(m.inspect_document(out)['comments'][0]['text'], '核对样本\nCheck sample')
        self.assertEqual(len(r['comment_ids']), 1)

    def test_existing_comment_content_and_anchor_preserved(self):
        comments = f'<w:comments xmlns:w="{m.W}"><w:comment w:id="7" w:author="Teacher"><w:p>{run("old comment")}</w:p></w:comment></w:comments>'
        body = '<w:p><w:commentRangeStart w:id="7"/>'+run('30次')+'<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>'+run(' 额外内容')+'</w:p>'
        p = package(body, {m.COMMENTS: comments.encode()})
        out, _ = self.build(p, plan(p, replace('30', '10'), {'type':'comment','paragraph':'p000001','quote':'额外内容','text':'new comment'}))
        self.assertEqual(len(m.inspect_document(out)['comments']), 2)
        self.assertIn('7', out.root.xpath('//w:commentRangeStart/@w:id', namespaces=m.NS))

    def test_conflicts_report_all_and_no_outputs(self):
        p = package('<w:p>' + run('same same') + '</w:p>')
        with tempfile.TemporaryDirectory() as tmp:
            src, dst, receipt = [Path(tmp)/x for x in ('in.docx','out.docx','receipt.json')]
            src.write_bytes(p.data)
            with self.assertRaises(m.Conflict) as cm:
                m.apply_file(src, plan(p, replace('same','new'), replace('missing','new')), dst, receipt)
            self.assertEqual(len(cm.exception.details), 2)
            self.assertFalse(dst.exists())
            self.assertFalse(receipt.exists())
            self.assertEqual(src.read_bytes(), p.data)

    def test_stale_source_and_overlapping_operations(self):
        p = package('<w:p>' + run('abcdef') + '</w:p>')
        stale = plan(p, replace('abc','ABC'));stale['source_sha256']='0'*64
        for candidate in (stale, plan(p, replace('abc','ABC'), replace('bcd','BCD'))):
            with self.assertRaises(m.Conflict):m.build(p, candidate)

    def test_fields_and_bookmark_crossing_rejected(self):
        for body, old in [
            (run('a')+'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'+run('b'), 'a'),
            (run('a')+'<w:bookmarkStart w:id="1" w:name="x"/>'+run('b')+'<w:bookmarkEnd w:id="1"/>', 'ab'),
            ('<w:hyperlink>'+run('a')+'</w:hyperlink>', 'a')]:
            p = package('<w:p>' + body + '</w:p>')
            with self.assertRaises(m.Conflict):m.build(p, plan(p, replace(old,'x')))

    def test_unchanged_drawing_and_hyperlink_outside_span(self):
        p = package('<w:p>'+run('30')+'<w:hyperlink>'+run('link')+'</w:hyperlink><w:r><w:drawing/></w:r></w:p>')
        out, _ = self.build(p, plan(p, replace('30','10')))
        self.assertEqual(m.canon(p.root.find('.//' + m.q('hyperlink'))), m.canon(out.root.find('.//' + m.q('hyperlink'))))

    def test_source_change_during_apply(self):
        p = package('<w:p>'+run('30')+'</w:p>')
        with tempfile.TemporaryDirectory() as tmp:
            src, dst, receipt = [Path(tmp)/x for x in ('in.docx','out.docx','receipt.json')]
            src.write_bytes(p.data)
            real = m.build
            def racing(*args, **kwargs):
                result = real(*args, **kwargs)
                src.write_bytes(b'concurrent change')
                return result
            with patch.object(m,'build',side_effect=racing):
                with self.assertRaises(m.Conflict):m.apply_file(src,plan(p,replace('30','10')),dst,receipt)
            self.assertFalse(dst.exists());self.assertFalse(receipt.exists())

    def test_apply_new_path_and_receipt(self):
        p = package('<w:p>'+run('30')+'</w:p>')
        with tempfile.TemporaryDirectory() as tmp:
            src,dst,receipt = [Path(tmp)/x for x in ('in.docx','out.docx','receipt.json')]
            src.write_bytes(p.data)
            pplan=plan(p,replace('30','10'))
            r=m.apply_file(src,pplan,dst,receipt)
            self.assertEqual(src.read_bytes(),p.data)
            self.assertTrue(m.verify(p,m.Package.read(dst),pplan,json.loads(receipt.read_text()))['ok'])
            with self.assertRaises(m.Conflict):m.apply_file(src,pplan,dst,receipt)
            self.assertEqual(m.sha(dst.read_bytes()),r['output_sha256'])

    def test_tampering_detected(self):
        p = package('<w:p>'+run('30')+'</w:p>')
        pplan=plan(p,replace('30','10'))
        out,r=self.build(p,pplan)
        altered = m.Package(out.pack({'word/media/image1.bin': b'changed'}))
        r['output_sha256']=altered.digest
        with self.assertRaises(m.Conflict):m.verify(p,altered,pplan,r)

    def test_tracking_false_and_settings_order(self):
        settings=f'<w:settings xmlns:w="{m.W}"><w:zoom w:percent="100"/><w:trackRevisions w:val="false"/><w:compat/></w:settings>'
        p=package('<w:p>'+run('30')+'</w:p>',settings=settings)
        out,_=self.build(p,plan(p,replace('30','10')))
        self.assertEqual(out.roots[m.SETTINGS].find(m.q('trackRevisions')).get(m.q('val')),'true')

    def test_custom_namespace_retained(self):
        p=package('<w:p>'+run('30')+'</w:p>')
        text=p.parts[m.MAIN].replace(b'<w:document ', b'<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14" ')
        p=m.Package(p.pack({m.MAIN:text}))
        out,_=self.build(p,plan(p,replace('30','10')))
        self.assertEqual(out.root.nsmap['w14'],p.root.nsmap['w14'])

    def test_comment_on_the_same_replacement(self):
        for new in ('10次评分', ''):
            with self.subTest(new=new):
                p=package('<w:p>'+run('重复30次评分。')+'</w:p>')
                out,r=self.build(p,plan(p,replace('30次评分',new,comment='Corrected repetition count')))
                self.assertEqual(m.text_view(out.paragraphs[0]),'重复'+new+'。')
                self.assertEqual(len(r['comment_ids']),1)
                self.assertEqual(m.inspect_document(out)['comments'][0]['text'],'Corrected repetition count')

    def test_unknown_operation_field_reported(self):
        p=package('<w:p>'+run('30')+'</w:p>')
        with self.assertRaises(m.Conflict):m.build(p,plan(p,replace('30','10',auther='Someone')))

    def test_comment_and_replace_on_separate_spans(self):
        p=package('<w:p>'+run('abc def ghi')+'</w:p>')
        out,_=self.build(p,plan(p,replace('abc','xyz'), {'type':'comment','paragraph':'p000001','quote':'ghi','text':'Check this'}))
        self.assertEqual(m.text_view(out.paragraphs[0]),'xyz def ghi')

    def test_cli_smoke(self):
        p=package('<w:p>'+run('30')+'</w:p>')
        with tempfile.TemporaryDirectory() as tmp:
            src=Path(tmp)/'in.docx';src.write_bytes(p.data)
            result=subprocess.run([sys.executable,str(SCRIPT),'inspect',str(src)],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertEqual(json.loads(result.stdout)['source_sha256'],p.digest)

    def test_malformed_receipt_cli_returns_json_error(self):
        p=package('<w:p>'+run('30')+'</w:p>')
        pplan=plan(p,replace('30','10'));out,r=self.build(p,pplan)
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)
            (root/'in.docx').write_bytes(p.data);(root/'out.docx').write_bytes(out.data)
            (root/'plan.json').write_text(json.dumps(pplan));(root/'receipt.json').write_text('[]')
            result=subprocess.run([sys.executable,str(SCRIPT),'verify',str(root/'in.docx'),str(root/'out.docx'),'--plan',str(root/'plan.json'),'--receipt',str(root/'receipt.json')],capture_output=True,text=True)
            self.assertEqual(result.returncode,2,result.stderr)
            self.assertFalse(json.loads(result.stdout)['ok'])

    def test_invalid_inputs(self):
        with self.assertRaises(m.Conflict):m.Package(b'not a zip')
        p=package('<w:p>'+run('30')+'</w:p>')
        for candidate in [plan(p, replace('30','x\x01')),plan(p,replace('30','line\nbreak')),plan(p,replace('30','30'))]:
            with self.assertRaises(m.Conflict):m.build(p,candidate)

    def test_move_range_spanning_paragraphs_rejected(self):
        p=package('<w:moveToRangeStart w:id="8" w:name="move"/><w:p>'+run('30')+'</w:p><w:moveToRangeEnd w:id="8"/>')
        with self.assertRaises(m.Conflict):m.build(p,plan(p,replace('30','10')))

    def test_structural_revision_rejected(self):
        p=package('<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Teacher"/></w:rPr></w:pPr>'+run('30')+'</w:p>')
        with self.assertRaises(m.Conflict):m.build(p,plan(p,replace('30','10')))


if __name__=='__main__':
    unittest.main()
