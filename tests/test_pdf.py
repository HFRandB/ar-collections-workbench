import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from fdl_pdf import build_pdf_filename


def test_filename_0038():
    name = build_pdf_filename('0038', '1302690', 'Tianneng Battery Group Co., Ltd.', '天能电池集团股份有限公司')
    assert name == '最终催款通知 - 1302690 Tianneng Battery Group Co., Ltd. 天能电池集团股份有限公司 - CN 0038.pdf'


def test_filename_0073():
    name = build_pdf_filename('0073', '195218', 'TSMC', '台灣積體電路製造股份有限公司')
    assert name == '最终催款通知 - 195218 TSMC 台灣積體電路製造股份有限公司 - TW 0073.pdf'


def test_filename_0074():
    name = build_pdf_filename('0074', '3087470', 'Cheerplan (HK) Limited', '')
    assert name == 'Final Demand Notice - 3087470 Cheerplan (HK) Limited - HK 0074.pdf'


def test_filename_0074_with_cn_name():
    name = build_pdf_filename('0074', '3087470', 'Cheerplan (HK) Limited', '吉祥計劃（香港）有限公司')
    assert name == 'Final Demand Notice - 3087470 Cheerplan (HK) Limited 吉祥計劃（香港）有限公司 - HK 0074.pdf'
