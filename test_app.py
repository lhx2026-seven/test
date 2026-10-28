import unittest
from io import BytesIO
from pptx import Presentation
import pandas as pd
from app import read_report, prepare, generate, ROOT

class WorkflowTests(unittest.TestCase):
    def test_csv_and_all_slides(self):
        frame=read_report('sample.csv',(ROOT/'sample.csv').read_bytes())
        data=prepare(frame,'品名','銷量','品類')
        self.assertEqual(data['銷量'].sum(),2205)
        for mode in ('銷量排行','品類占比','低銷量觀察'):
            output=generate(data,mode,'測試週報','2026/10/19–25')
            prs=Presentation(BytesIO(output));self.assertEqual(len(prs.slides),1)
            charts=[s.chart for s in prs.slides[0].shapes if s.has_chart]
            self.assertEqual(len(charts),1)
            self.assertTrue(any(s.has_text_frame and '2,205' in s.text for s in prs.slides[0].shapes))
            self.assertGreater(len(charts[0].series[0].values),0)
    def test_excel_and_aggregation(self):
        frame=pd.DataFrame({'品名':['A','A','B'],'銷量':['1,000','20','5']})
        buf=BytesIO();frame.to_excel(buf,index=False)
        data=prepare(read_report('test.xlsx',buf.getvalue()),'品名','銷量',None)
        self.assertEqual(data.groupby('品項')['銷量'].sum()['A'],1020)
    def test_invalid_quantities(self):
        for values in [['x'],[-1],[0],['']]:
            with self.assertRaises(ValueError): prepare(pd.DataFrame({'p':['A'],'q':values}),'p','q',None)
    def test_empty_product(self):
        with self.assertRaises(ValueError):prepare(pd.DataFrame({'p':[' '],'q':[2]}),'p','q',None)

if __name__=='__main__':unittest.main()
