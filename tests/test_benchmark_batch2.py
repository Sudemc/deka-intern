"""Benchmark tests for Batch 2: Acoustic (BUZ, 15 items) and Semiconductor Passives (DIO, 14 items)."""

import pytest
from bom_enricher.datasheet_analyzer import analyze_datasheet
from bom_enricher.comparison import load_rulebook
from pathlib import Path


@pytest.fixture(scope="module")
def rulebook():
    return load_rulebook(Path("data/rules.xlsx"))


def test_batch2_buz_buzzers(rulebook):
    # 1. Murata PKLCS1212E40A1-R1
    res = analyze_datasheet(b"", part_number="PKLCS1212E40A1-R1", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZSA030307-MT"
    assert "Murata" in res["rich_description"]

    # 2. FUET FUET-5020
    res = analyze_datasheet(b"", part_number="FUET-5020", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZSA030310-FT"
    assert "FUET" in res["rich_description"]

    # 3. TRUWATT (FUET-5020)
    res = analyze_datasheet(b"", part_number="TRUWATT (FUET-5020)", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZSA030310-TRU"
    assert "TRUWATT" in res["rich_description"]

    # 4. CORE SMT-0540-T-9-R
    res = analyze_datasheet(b"", part_number="SMT-0540-T-9-R", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZSA030411-COR"

    # 5. YXDZ YX7525P
    res = analyze_datasheet(b"", part_number="YX7525P", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZSA060608-YXZ"
    assert "YXDZ" in res["rich_description"]

    # 6. Be Star BPT1703H09-02
    res = analyze_datasheet(b"", part_number="BPT1703H09-02", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTA010205-BS"

    # 7. Kepo KPT-G1713-K11236
    res = analyze_datasheet(b"", part_number="KPT-G1713-K11236", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTA010205-KE"

    # 8. Kepo KPT-G2312-K7462
    res = analyze_datasheet(b"", part_number="KPT-G2312-K7462", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTA010206-KE"

    # 9. SANCO EPT-22A2003P
    res = analyze_datasheet(b"", part_number="EPT-22A2003P", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTA010206-SAN"

    # 10. SOUND COMP. PT2770AHP-06
    res = analyze_datasheet(b"", part_number="PT2770AHP-06", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTA010206-SOU"

    # 11. YXDZ YX-YT9503
    res = analyze_datasheet(b"", part_number="YX-YT9503", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTA100400-YXZ"

    # 12. YXDZ YX-YT12065-YT05D
    res = analyze_datasheet(b"", part_number="YX-YT12065-YT05D", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTB010602-YXZ"

    # 13. Hong De Da Elec. HPT-1780P-05TB-2.0-30-10.0
    res = analyze_datasheet(b"", part_number="HPT-1780P-05TB-2.0-30-10.0", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTB010605-HD"

    # 14. Huaneng HNB09A05
    res = analyze_datasheet(b"", part_number="HNB09A05", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTB020600-HUA"

    # 15. GENTEK YX-WY1470A
    res = analyze_datasheet(b"", part_number="GENTEK (YX-WY1470A)", rulebook=rulebook)
    assert res["material_type"] == "BUZ"
    assert res["erp_code"] == "BUZTB030604-GEN"


def test_batch2_dio_diodes(rulebook):
    # 1. Philips/Nxp BZV55-C20,115
    res = analyze_datasheet(b"", part_number="BZV55-C20,115", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSA13H900-PH"

    # 2. Nexperia PMEG2010ER,115
    res = analyze_datasheet(b"", part_number="PMEG2010ER,115", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSD81I1300-NP"

    # 3. On Semi BAS19LT1G
    res = analyze_datasheet(b"", part_number="BAS19LT1G", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSSA0J000-ON"

    # 4. Panjit P6SMBJ120CA(S)_R1_00001
    res = analyze_datasheet(b"", part_number="P6SMBJ120CA(S)_R1_00001", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOST56K900-PA"

    # 5. Panjit P6SMBJ120CA_R3_00001
    res = analyze_datasheet(b"", part_number="P6SMBJ120CA_R3_00001", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOST56M1900-PA"

    # 6. TRUWATT (SMBJ20A)
    res = analyze_datasheet(b"", part_number="TRUWATT (SMBJ20A)", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOST56Z2600-TRU"

    # 7. ANOVA (Onsemi) NSR0520V2T1G
    res = analyze_datasheet(b"", part_number="560232-001 (NSR0520V2T1G)", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSV81V500-AOV"

    # 8. LRC LMBR0520T1G
    res = analyze_datasheet(b"", part_number="LMBR0520T1G", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSZ101Z500-L0"

    # 9. Nexperia NZH20C,115 (Automotive AEC-Q101)
    res = analyze_datasheet(b"", part_number="NZH20C,115", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSZ103E17AC-NP"

    # 10. Hottech MBR0520-HT
    res = analyze_datasheet(b"", part_number="MBR0520-HT", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSZ41M600-HO"

    # 11. Panjit SS0520_R1_00001
    res = analyze_datasheet(b"", part_number="SS0520_R1_00001", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSZ41S1000-PA"

    # 12. DTSİS (UMW) PSR05-LF-T7
    res = analyze_datasheet(b"", part_number="H010101005318 (PSR05-LF-T7)", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOSZZ96Z7100-DTS"

    # 13. Littelfuse SMAJ30CA
    res = analyze_datasheet(b"", part_number="SMAJ30CA", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOST46B1300-LT"

    # 14. Yageo SMAJ30CA-AT/TR13 (Automotive AEC-Q101)
    res = analyze_datasheet(b"", part_number="SMAJ30CA-AT/TR13", rulebook=rulebook)
    assert res["material_type"] == "DIO"
    assert res["erp_code"] == "DIOST46B13AC-YG"
