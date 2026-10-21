"""Benchmark test suite for Batch 5: Interconnects (CON - 39 Connectors, Headers, Terminals)."""

from pathlib import Path
import pytest
from bom_enricher.comparison import load_rulebook
from bom_enricher.datasheet_analyzer import analyze_datasheet


@pytest.fixture(scope="module")
def rulebook():
    return load_rulebook(Path("data/rules.xlsx"))


def test_batch5_con_connectors(rulebook):
    # 1. CONCAIB20300006-TL
    res = analyze_datasheet(b"", text="Cable Wire to Wire Klemens, Dişi, 8mm, 3 pin, Tianli  TLB-100-03P", rulebook=rulebook, serial_counter="0006")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONCAIB20300006-TL"

    # 2. CONCAIB40300113-GA
    res = analyze_datasheet(b"", text="Cable Wire to Wire Klemens, Dişi, 8.2mm, 3 pin, Gaosong/Degson    DG8HS-1.2-03P-17-100SH/AH", rulebook=rulebook, serial_counter="0113")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONCAIB40300113-GA"

    # 3. CONCAIB41200084-GA
    res = analyze_datasheet(b"", text="Cable Wire to Wire Klemens, Dişi, 8.2mm, 12 pin, Gaosong/Degson   DG8HS-1.2-12P-17-100SH/AH", rulebook=rulebook, serial_counter="0084")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONCAIB41200084-GA"

    # 4. CONSAABB0210785-ISM
    res = analyze_datasheet(b"", text="SMD Pin Header, Dişi, 2mm, 2 pin, 90°(Yatık), İsmak (Kls)   03030520017-00 (L-KLS1-208BC-6.35-1-2TWR", rulebook=rulebook, serial_counter="0785")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAABB0210785-ISM"

    # 5. CONSAABD0E20953-ENV
    res = analyze_datasheet(b"", text="SMD Pin Header, Dişi, 2.54mm, 2x6 pin, 180°(Dik), Enervon (Connfly)     HM00900002800 (DS1023-45-2X6B8SRS)", rulebook=rulebook, serial_counter="0953")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAABD0E20953-ENV"

    # 6. CONSABBC0210335-SC
    res = analyze_datasheet(b"", text="SMD IDC Konnektör, Dişi, 2.5mm, 2 pin, 90°(Yatık),  Stocko  CON-MFMP-7238-202-061-960-G", rulebook=rulebook, serial_counter="0335")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSABBC0210335-SC"

    # 7. CONSADBF0310940-WA
    res = analyze_datasheet(b"", text="SMD Wire to Board Konnektör, Dişi, 4mm, 3 pin, 90°(Yatık),  Wago  2060-453/998-404", rulebook=rulebook, serial_counter="0940")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSADBF0310940-WA"

    # 8. CONSAEBL0311112-DTS
    res = analyze_datasheet(b"", text="SMD Wire to Board Kilitli Konnektör, Dişi, 1.25mm, 3 pin, 90°(Yatık), DTSİS (JST)   H010202004601 (455-1565-1-ND)", rulebook=rulebook, serial_counter="1112")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAEBL0311112-DTS"

    # 9. CONSAEBL0811111-DTS
    res = analyze_datasheet(b"", text="SMD Wire to Board Kilitli Konnektör, Dişi, 1.25mm, 8 pin, 90°(Yatık), DTSİS (JST)   H010202004606 (455-1570-1-ND)", rulebook=rulebook, serial_counter="1111")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAEBL0811111-DTS"

    # 10. CONSAG0B00621037-AM
    res = analyze_datasheet(b"", text="SMD Board to Board Konnektör, Dişi, 6 pin, 180°(Dik), Amphenol ICC      89898-303ALF", rulebook=rulebook, serial_counter="1037")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAG0B00621037-AM"

    # 11. CONSAHB12510001-MO
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.3mm, 35 pin, 90°(Yatık), Molex     502250-3591", rulebook=rulebook, serial_counter="0001")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB12510001-MO"

    # 12. CONSAHB12510243-MO
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.3mm, 35 pin, 90°(Yatık),  Molex    5047543520", rulebook=rulebook, serial_counter="0243")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB12510243-MO"

    # 13. CONSAHB12510307-OP
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.3mm, 35 pin, 90°(Yatık), Oupiin    2566-F35G00DT", rulebook=rulebook, serial_counter="0307")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB12510307-OP"

    # 14. CONSAHB12510909-HRE
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.3mm, 35 pin, 90°(Yatık),  Hirose Electric      FH35C-35S-0.3SHW(50)", rulebook=rulebook, serial_counter="0909")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB12510909-HRE"

    # 15. CONSAHB50610561-ZZ6
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 6 pin, 90°(Yatık), GCT  FFC2B35-06-G", rulebook=rulebook, serial_counter="0561")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB50610561-ZZ6"

    # 16. CONSAHB51011085-AM
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 10 pin, 90°(Yatık), Amphenol ICC    SFV10R-2STBE1HLF", rulebook=rulebook, serial_counter="1085")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB51011085-AM"

    # 17. CONSAHB51611075-COR
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 16 pin, 90°(Yatık), CORE (GCT)      CR-HM-004-0040(FFC2B28-16-G)", rulebook=rulebook, serial_counter="1075")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB51611075-COR"

    # 18. CONSAHB52011014-XUN
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 20 pin, 90°(Yatık),  XUNPU    FPC-05F-20PH20", rulebook=rulebook, serial_counter="1014")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB52011014-XUN"

    # 19. CONSAHB52310560-AM
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 40 pin, 90°(Yatık), Amphenol ICC    62684-402100ALF", rulebook=rulebook, serial_counter="0560")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB52310560-AM"

    # 20. CONSAHB52311110-DTS
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 40 pin, 90°(Yatık), DTSİS (Amphenol FCI)  H010202004302 (62684-402100ALF)", rulebook=rulebook, serial_counter="1110")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB52311110-DTS"

    # 21. CONSAHB53010112-WU
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 30 pin, 90°(Yatık), Wurth     68713014022", rulebook=rulebook, serial_counter="0112")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB53010112-WU"

    # 22. CONSAHB53010553-MO
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 30 pin, 90°(Yatık),  Molex    0541323062", rulebook=rulebook, serial_counter="0553")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB53010553-MO"

    # 23. CONSAHB53010559-MO
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 30 pin, 90°(Yatık), Molex     5051103091", rulebook=rulebook, serial_counter="0559")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB53010559-MO"

    # 24. CONSAHB5301100-MO
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 30 pin, 90°(Yatık), Molex     0541043033", rulebook=rulebook, serial_counter="100")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB5301100-MO"

    # 25. CONSAHB53420505-AM
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.5mm, 67 pin, 180°(Dik),  Amphenol ICC    10131758-001RLF", rulebook=rulebook, serial_counter="0505")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHB53420505-AM"

    # 26. CONSAHBA0510002-WU
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 1mm, 5 pin, 90°(Yatık), Wurth  68610514422", rulebook=rulebook, serial_counter="0002")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHBA0510002-WU"

    # 27. CONSAHBA051099-MO
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 1mm, 5 pin, 90°(Yatık), Molex  0522070560", rulebook=rulebook, serial_counter="099")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHBA051099-MO"

    # 28. CONSAHBA1210306-MO
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 1mm, 12 pin, 90°(Yatık), Molex 52207-1233", rulebook=rulebook, serial_counter="0306")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHBA1210306-MO"

    # 29. CONSAHBZ00811073-COR
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.50mm, 8 pin, 90°(Yatık), CORE (Amphenol ICC)   CR-HM-004-0024 (F3311A7H121008E200)", rulebook=rulebook, serial_counter="1073")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHBZ00811073-COR"

    # 30. CONSAHBZ00811074-COR
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.50mm, 8 pin, 90°(Yatık), CORE (GCT)      CR-HM-004-0024 (FFC2B28-08-G)", rulebook=rulebook, serial_counter="1074")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHBZ00811074-COR"

    # 31. CONSAHBZ01011047-AM
    res = analyze_datasheet(b"", text="SMD FFC/FPC Konnektör, Dişi, 0.50mm, 10 pin, 90°(Yatık), Amphenol ICC   SFV10R-1STBE1HLF", rulebook=rulebook, serial_counter="1047")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAHBZ01011047-AM"

    # 32. CONSAJB51010302-IR
    res = analyze_datasheet(b"", text="SMD Box Header, Dişi, 0.5mm, 10 pin (2x5), 90°(Yatık), (IMSA-10112B-10A), IRISO     IMSA-10112B-10Y900", rulebook=rulebook, serial_counter="0302")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAJB51010302-IR"

    # 33. CONSAJBD0D20239-WU
    res = analyze_datasheet(b"", text="SMD Box Header, Dişi, 2.54mm, 2x5 pin, 180°(Dik),  Wurth    690367291076", rulebook=rulebook, serial_counter="0239")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAJBD0D20239-WU"

    # 34. CONSAJBE1010301-MO
    res = analyze_datasheet(b"", text="SMD Box Header, Dişi, 0.80mm, 10 pin (2x5), 90°(Yatık), Molex     0533091070", rulebook=rulebook, serial_counter="0301")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAJBE1010301-MO"

    # 35. CONSANB00511038-MIK
    res = analyze_datasheet(b"", text="SMD USB-microB, Dişi, 5 pin, 90°(Yatık), Mikrodev (On Shore)      HM0064 (USB-M26FTR)", rulebook=rulebook, serial_counter="1038")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSANB00511038-MIK"

    # 36. CONSAOB02A10546-MO
    res = analyze_datasheet(b"", text="SMD Modular Jack, Dişi,  6P6C, 90°(Yatık),  (ELDOR) Molex   0437436001", rulebook=rulebook, serial_counter="0546")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAOB02A10546-MO"

    # 37. CONSAPBD0420320-WU
    res = analyze_datasheet(b"", text="SMD Socket, Dişi, 2.54mm, 4 pin, 180°(Dik),  Wurth    61000413321", rulebook=rulebook, serial_counter="0320")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSAPBD0420320-WU"

    # 38. CONSASBE0420308-KL
    res = analyze_datasheet(b"", text="TH, USB MINI B Type, Dişi, 0.80mm, 4 pin, 180°(Dik),  Kls   L-KLS1-229-5FA-B", rulebook=rulebook, serial_counter="0308")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSASBE0420308-KL"

    # 39. CONSATBO1020813-ZZ3
    res = analyze_datasheet(b"", text="SMD Receptacles, Dişi, 1.27mm, 10 pin, 180°(Dik),  Samtec   CLP-105-02-L-D-TR", rulebook=rulebook, serial_counter="0813")
    assert res["material_type"] == "CON"
    assert res["erp_code"] == "CONSATBO1020813-ZZ3"

