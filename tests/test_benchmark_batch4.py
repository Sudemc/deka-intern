"""Benchmark test suite for Batch 4: Frequency Control (CRY - 66 Crystals, Oscillators, Resonators)."""

from pathlib import Path
import pytest
from bom_enricher.comparison import load_rulebook
from bom_enricher.datasheet_analyzer import analyze_datasheet


@pytest.fixture(scope="module")
def rulebook():
    return load_rulebook(Path("data/rules.xlsx"))


def test_batch4_cry_crystals(rulebook):
    # 1. CRYSA0078-EPS | FC-135 32.7680KA-A5
    res = analyze_datasheet(b"", text="SMD 2-SMD, 32.768kHz, 12.5pF, 70K Ohms, EPSON", part_number="FC-135 32.7680KA-A5", rulebook=rulebook, serial_counter="078")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0078-EPS"

    # 2. CRYSA0079-ZZ9 | ECS-.327-12.5-12-TR
    res = analyze_datasheet(b"", text="SMD 2-SMD, 32.768kHz, 12.5pF, 90kOhms, ECS", part_number="ECS-.327-12.5-12-TR", rulebook=rulebook, serial_counter="079")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0079-ZZ9"

    # 3. CRYSA0081-GEN | GENTEK (FC31M2-32.768-NTLLLDT)
    res = analyze_datasheet(b"", text="SMD SMD3215-2P, 32.768kHz, 12.5pF 50kOhm, GENTEK (HCI)", part_number="GENTEK (FC31M2-32.768-NTLLLDT)", rulebook=rulebook, serial_counter="081")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0081-GEN"

    # 4. CRYSA0084-ARC | ABS06-107-32.768KHZ-T
    res = analyze_datasheet(b"", text="SMD 2-SMD, 32.768kHz, 4pF, 60kOhms, Abracon", part_number="ABS06-107-32.768KHZ-T", rulebook=rulebook, serial_counter="084")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0084-ARC"

    # 5. CRYSA0085-ENV | HM01400000200 (X1A0000610002)
    res = analyze_datasheet(b"", text="SMD SMD2012-2P, 32.768kHz, 12.5pF, 90kOhm, Enervon (EPSON)", part_number="HM01400000200 (X1A0000610002)", rulebook=rulebook, serial_counter="085")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0085-ENV"

    # 6. CRYSA0090-HCI | FC31M2-32.768-NTLLLDT
    res = analyze_datasheet(b"", text="SMD SMD3215-2P, 32.768kHz, 12.5pF 50kOhm, HCI", part_number="FC31M2-32.768-NTLLLDT", rulebook=rulebook, serial_counter="090")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0090-HCI"

    # 7. CRYSA0094-SKI | SC20S-7PF20PPM
    res = analyze_datasheet(b"", text="SMD 2-SMD, 32.768kHz, 7pF, 70kOhms, Seiko Instruments", part_number="SC20S-7PF20PPM", rulebook=rulebook, serial_counter="094")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0094-SKI"

    # 8. CRYSA0098-RA | RT3215-32.768-12.5-TR
    res = analyze_datasheet(b"", text="SMD 32.768kHz, 70 kOhms, 12.5P, 20PPM, Raltron", part_number="RT3215-32.768-12.5-TR", rulebook=rulebook, serial_counter="098")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0098-RA"

    # 9. CRYSA0099-MIK | HM0537 (MC-306 32.768000KHZ)
    res = analyze_datasheet(b"", text="SMD 4-SOJ, 12.5pF, 32.768kHz, Mikrodev (Epson)", part_number="HM0537 (MC-306 32.768000KHZ)", rulebook=rulebook, serial_counter="099")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA0099-MIK"

    # 10. CRYSA2022-ZZ9 | ECS-80-20-5PX-TR
    res = analyze_datasheet(b"", text="SMD 8MHz, ECS", part_number="ECS-80-20-5PX-TR", rulebook=rulebook, serial_counter="022")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA2022-ZZ9"

    # 11. CRYSA2025-HDE | HDA001646(TX50328M2KDCE1T)
    res = analyze_datasheet(b"", text="SMD Crystal Oscillator, 8MHz, HDAEnerji (Yajingxin)", part_number="HDA001646(TX50328M2KDCE1T)", rulebook=rulebook, serial_counter="025")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA2025-HDE"

    # 12. CRYSA2028-RA | H130B-8.000-18-3030-EXT
    res = analyze_datasheet(b"", text="SMD 8MHz, 18pF, H130BR, Raltron", part_number="H130B-8.000-18-3030-EXT", rulebook=rulebook, serial_counter="028")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA2028-RA"

    # 13. CRYSA2042-MT | CSTNE8M00G52A000R0
    res = analyze_datasheet(b"", text="SMD 8MHz,  Murata", part_number="CSTNE8M00G52A000R0", rulebook=rulebook, serial_counter="042")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA2042-MT"

    # 14. CRYSA2076-RA | RH100-8.000-16-3030-TR-N
    res = analyze_datasheet(b"", text="SMD 4-SMD, 8MHz, 16 pF, 400Ohm, Raltron", part_number="RH100-8.000-16-3030-TR-N", rulebook=rulebook, serial_counter="076")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA2076-RA"

    # 15. CRYSA2091-TRU | TRUWATT (TAXM8M4RDBCCT2T)
    res = analyze_datasheet(b"", text="SMD Crystal Oscillator, 8MHz, 10pF, 4P, TRUWATT (YAJINGXIN)", part_number="TRUWATT (TAXM8M4RDBCCT2T)", rulebook=rulebook, serial_counter="091")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA2091-TRU"

    # 16. CRYSA2093-TRU | TRUWATT (306010014200)
    res = analyze_datasheet(b"", text="SMD SMD3225, 8MHz, 12pF, TRUWATT", part_number="TRUWATT (306010014200)", rulebook=rulebook, serial_counter="093")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA2093-TRU"

    # 17. CRYSA4050-KY0 | CX3225GA20000D0PTVCC
    res = analyze_datasheet(b"", text="SMD 2-SMD, 20MHz, 8pF, 100 Ohms, AEC-Q200, KYOCERA AVX", part_number="CX3225GA20000D0PTVCC", rulebook=rulebook, serial_counter="050")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA4050-KY0"

    # 18. CRYSA5001-JA | Q 26,0-JXS32-10-10/10-FU-LF
    res = analyze_datasheet(b"", text="SMD 26MHz, Jauch", part_number="Q 26,0-JXS32-10-10/10-FU-LF", rulebook=rulebook, serial_counter="001")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA5001-JA"

    # 19. CRYSA5061-EPS | TSX-3225 26.0000MF09Z-AC3
    res = analyze_datasheet(b"", text="SMD 4-SMD, 26MHz, 9pF, 40Ohm, EPSON", part_number="TSX-3225 26.0000MF09Z-AC3", rulebook=rulebook, serial_counter="061")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA5061-EPS"

    # 20. CRYSA6089-MT | CSTCR4M00G53Z-R0
    res = analyze_datasheet(b"", text="SMD 3-SMD, Non-Standard, 4MHz, 15 pF,  60 Ohms, Murata", part_number="CSTCR4M00G53Z-R0", rulebook=rulebook, serial_counter="089")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA6089-MT"

    # 21. CRYSA7027-SST | SST 13.560M 20PF 10PPM
    res = analyze_datasheet(b"", text="SMD 13.560MHz, Source-Set", part_number="SST 13.560M 20PF 10PPM", rulebook=rulebook, serial_counter="027")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA7027-SST"

    # 22. CRYSA9009-DNI | DANNİE (NX3225SA-12MHZ-STD-CSR-6)
    res = analyze_datasheet(b"", text="SMD 12MHz, Dannie.CC (NDK)", part_number="DANNİE (NX3225SA-12MHZ-STD-CSR-6)", rulebook=rulebook, serial_counter="009")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA9009-DNI"

    # 23. CRYSA9012-RA | RH100-12.000-12-1010-EXT
    res = analyze_datasheet(b"", text="SMD 12MHz, 3.2x2.5mm, Raltron", part_number="RH100-12.000-12-1010-EXT", rulebook=rulebook, serial_counter="012")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA9012-RA"

    # 24. CRYSA9020-ARC | ABM8G-12.000MHZ-4Y-T3
    res = analyze_datasheet(b"", text="SMD 12MHz, 10pF, ABM8G,  Abracon", part_number="ABM8G-12.000MHZ-4Y-T3", rulebook=rulebook, serial_counter="020")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA9020-ARC"

    # 25. CRYSA9102-DTS | H010102012143 (3TJ412000HYFBC)
    res = analyze_datasheet(b"", text="SMD 3225-4P, 12MHz, 9pF, ±10ppm, DTSİS (JYJE)", part_number="H010102012143 (3TJ412000HYFBC)", rulebook=rulebook, serial_counter="102")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSA9102-DTS"

    # 26. CRYSB0008-DNI | DANNİE (NX2520SA-16MHZ-STD-CSW-5)
    res = analyze_datasheet(b"", text="SMD 16MHz, Dannie.CC (NDK)", part_number="DANNİE (NX2520SA-16MHZ-STD-CSW-5)", rulebook=rulebook, serial_counter="008")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0008-DNI"

    # 27. CRYSB0013-RA | RH100-16.000-10-1010-EXT
    res = analyze_datasheet(b"", text="SMD 16MHz, 3.2x2.5mm, Raltron", part_number="RH100-16.000-10-1010-EXT", rulebook=rulebook, serial_counter="013")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0013-RA"

    # 28. CRYSB0032-RA | RH100-16.000-10-1010-EXT-TR
    res = analyze_datasheet(b"", text="SMD 16MHz, 3.2x2.5mm, Raltron", part_number="RH100-16.000-10-1010-EXT-TR", rulebook=rulebook, serial_counter="032")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0032-RA"

    # 29. CRYSB0035-ZZ9 | ECS-163.84-18-5PLX-AGN-TR
    res = analyze_datasheet(b"", text="SMD 16MHz, 11.4x4.8mm, ECS", part_number="ECS-163.84-18-5PLX-AGN-TR", rulebook=rulebook, serial_counter="035")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0035-ZZ9"

    # 30. CRYSB0041-RA | AS-16.384-18-2050-SMD-TR
    res = analyze_datasheet(b"", text="SMD 16MHz, 4.7X4.6mm, Raltron", part_number="AS-16.384-18-2050-SMD-TR", rulebook=rulebook, serial_counter="041")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0041-RA"

    # 31. CRYSB0054-ARC | ABM10AIG-16.000MHZ-4Z-T3
    res = analyze_datasheet(b"", text="SMD 4-SMD, 16MHz, 10pF, 80 Ohms, AEC-Q200, Abracon", part_number="ABM10AIG-16.000MHZ-4Z-T3", rulebook=rulebook, serial_counter="054")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0054-ARC"

    # 32. CRYSB0055-MPL | M-PLUSTEC(ABM10AIG-16.000MHZ-4Z-T3)
    res = analyze_datasheet(b"", text="SMD 4-SMD, 16MHz, 10pF, 80 Ohms, AEC-Q200, M-Plustec (Abracon)", part_number="M-PLUSTEC(ABM10AIG-16.000MHZ-4Z-T3)", rulebook=rulebook, serial_counter="055")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0055-MPL"

    # 33. CRYSB0065-ENV | HM01400000300
    res = analyze_datasheet(b"", text="SMD 16MHz, 8pF, Enervon", part_number="HM01400000300", rulebook=rulebook, serial_counter="065")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0065-ENV"

    # 34. CRYSB0080-ENV | HM01400000300 (XRCGB16M000FXN20R0)
    res = analyze_datasheet(b"", text="SMD 4-SMD 16MHz, 8pF, 270Ohms, Enervon (Murata)", part_number="HM01400000300 (XRCGB16M000FXN20R0)", rulebook=rulebook, serial_counter="080")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0080-ENV"

    # 35. CRYSB0087-ENV | HM01400000300 (830108178109)
    res = analyze_datasheet(b"", text="SMD 4-SMD 16MHz, 8pF, 200Ohms, Enervon (Wurth)", part_number="HM01400000300 (830108178109)", rulebook=rulebook, serial_counter="087")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0087-ENV"

    # 36. CRYSB0100-COR | CR-HM-012-0004 (830055951)
    res = analyze_datasheet(b"", text="SMD 16MHz,18pF, 70Ohm, CORE (Würth)", part_number="CR-HM-012-0004 (830055951)", rulebook=rulebook, serial_counter="100")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0100-COR"

    # 37. CRYSB0103-DTS | H010102012072(TSX-3225 16.0000MF09Z-AC0)
    res = analyze_datasheet(b"", text="SMD 4-SMD, 16MHz, ±10ppm, 9pF, 60 Ohms, DTSİS (EPSON)", part_number="H010102012072(TSX-3225 16.0000MF09Z-AC0)", rulebook=rulebook, serial_counter="103")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB0103-DTS"

    # 38. CRYSB1014-EPT | FA-238 24.0000MB-C3
    res = analyze_datasheet(b"", text="SMD 24MHz, Epson Timing", part_number="FA-238 24.0000MB-C3", rulebook=rulebook, serial_counter="014")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB1014-EPT"

    # 39. CRYSB1019-ARC | ASDMB-24.000MHZ-LC-T
    res = analyze_datasheet(b"", text="SMD 24MHz, MEMS OSC, Abracon", part_number="ASDMB-24.000MHZ-LC-T", rulebook=rulebook, serial_counter="019")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB1019-ARC"

    # 40. CRYSB1048-MT | XRCGB24M000F2P91R0
    res = analyze_datasheet(b"", text="SMD 4-SMD, 24MHz,8pF, 60 Ohms, Murata", part_number="XRCGB24M000F2P91R0", rulebook=rulebook, serial_counter="048")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB1048-MT"

    # 41. CRYSB1096-MIK | HM0039 (ECS-240-6-37-JTN-TR)
    res = analyze_datasheet(b"", text="SMD 4-SMD, 24MHz, 6pF, Mikrodev (ECS)", part_number="HM0039 (ECS-240-6-37-JTN-TR)", rulebook=rulebook, serial_counter="096")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB1096-MIK"

    # 42. CRYSB1104-DTS | H010102012078(TSX-3225 24.0000MF15X-AC6)
    res = analyze_datasheet(b"", text="SMD 4-SMD, 24MHz, ±10ppm, 9pF, 40 Ohms, DTSİS (EPSON)", part_number="H010102012078(TSX-3225 24.0000MF15X-AC6)", rulebook=rulebook, serial_counter="104")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB1104-DTS"

    # 43. CRYSB2015-EPT | FA-238 25.0000MB-C3
    res = analyze_datasheet(b"", text="SMD 25MHz, Epson Timing", part_number="FA-238 25.0000MB-C3", rulebook=rulebook, serial_counter="015")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB2015-EPT"

    # 44. CRYSB2017-ARC | ASDMB-25.000MHZ-LC-T
    res = analyze_datasheet(b"", text="SMD 25MHz, MEMS OSC, Abracon", part_number="ASDMB-25.000MHZ-LC-T", rulebook=rulebook, serial_counter="017")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB2017-ARC"

    # 45. CRYSB2038-NDK | NX3225GA-25.000M-STD-CRG-2
    res = analyze_datasheet(b"", text="SMD 25MHz,  NDK", part_number="NX3225GA-25.000M-STD-CRG-2", rulebook=rulebook, serial_counter="038")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB2038-NDK"

    # 46. CRYSB2040-RA | RH100-25.000-18-F-3030-T
    res = analyze_datasheet(b"", text="SMD 25MHz, 3.2X2.5mm, Raltron", part_number="RH100-25.000-18-F-3030-T", rulebook=rulebook, serial_counter="040")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB2040-RA"

    # 47. CRYSB2059-RA | AS-25.000-20-F-3030-EXT-SMD-TR
    res = analyze_datasheet(b"", text="SMD HC-49/US, 25MHz, 20pF, 40Ohms, Raltron", part_number="AS-25.000-20-F-3030-EXT-SMD-TR", rulebook=rulebook, serial_counter="059")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB2059-RA"

    # 48. CRYSB2072-TGG | TOGG (XRCGB25M000F3A16R0)
    res = analyze_datasheet(b"", text="SMD 4-SMD 25MHz, 10pF, 80Ohms, AEC-Q200, Togg (Murata)", part_number="TOGG (XRCGB25M000F3A16R0)", rulebook=rulebook, serial_counter="072")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB2072-TGG"

    # 49. CRYSB3018-ARC | ASDMB-50.000MHZ-LC-T
    res = analyze_datasheet(b"", text="SMD 50MHz, MEMS OSC, Abracon", part_number="ASDMB-50.000MHZ-LC-T", rulebook=rulebook, serial_counter="018")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB3018-ARC"

    # 50. CRYSB3095-MIK | HM0032 (ASEMB-50.000MHZ-LC-T)
    res = analyze_datasheet(b"", text="SMD 4-SMD, 50MHz, 15pF, Mikrodev (Abracon)", part_number="HM0032 (ASEMB-50.000MHZ-LC-T)", rulebook=rulebook, serial_counter="095")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB3095-MIK"

    # 51. CRYSB4030-RA | RH100-40.000-18-F-2050-T
    res = analyze_datasheet(b"", text="SMD 40MHz, 18pF, Raltron", part_number="RH100-40.000-18-F-2050-T", rulebook=rulebook, serial_counter="030")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB4030-RA"

    # 52. CRYSB4053-ISM | 03030660001-00 (FW4000015Q)
    res = analyze_datasheet(b"", text="SMD 4-SMD 40MHz, AEC-Q200, İsmak (Diodes)", part_number="03030660001-00 (FW4000015Q)", rulebook=rulebook, serial_counter="053")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB4053-ISM"

    # 53. CRYSB4073-TGG | TOGG (NX2016SA-40M-STD-CZS-3)
    res = analyze_datasheet(b"", text="SMD 40MHz, 8pF, 2x1.6mm, AEC-Q200, Togg (NDK)", part_number="TOGG (NX2016SA-40M-STD-CZS-3)", rulebook=rulebook, serial_counter="073")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB4073-TGG"

    # 54. CRYSB4086-TGG | TOGG (STDCZS3-40M)
    res = analyze_datasheet(b"", text="SMD 4-SMD, 40MHz, 8PF, 60Ohms, AEC-Q200 Togg (NDK)", part_number="TOGG (STDCZS3-40M)", rulebook=rulebook, serial_counter="086")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB4086-TGG"

    # 55. CRYSB5058-TXC | 8Y39072005
    res = analyze_datasheet(b"", text="SMD 4-SMD, 39 MHz, 10pF, 35 Ohms, TXC Corporation", part_number="8Y39072005", rulebook=rulebook, serial_counter="058")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB5058-TXC"

    # 56. CRYSB6066-ZZ9 | ECS-270-12-33Q-JES-TR
    res = analyze_datasheet(b"", text="SMD 4-SMD, 27MHz, 12pF, 40Ohms, ECS", part_number="ECS-270-12-33Q-JES-TR", rulebook=rulebook, serial_counter="066")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB6066-ZZ9"

    # 57. CRYSB7083-GEN | GENTEK (XL7EL89COI-111YLC-32M)
    res = analyze_datasheet(b"", text="SMD SMD2016-4P, 32MHz, 12pF 100Ohm, GENTEK (YXC)", part_number="GENTEK (XL7EL89COI-111YLC-32M)", rulebook=rulebook, serial_counter="083")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB7083-GEN"

    # 58. CRYSB7088-ETR | EATRON (CX2016SA32000D0WRQG1)
    res = analyze_datasheet(b"", text="SMD 4-SMD 32MHz, 8pF, 60 Ohms, AEC-Q200, EATRON (KYOCERA)", part_number="EATRON (CX2016SA32000D0WRQG1)", rulebook=rulebook, serial_counter="088")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB7088-ETR"

    # 59. CRYSB8105-DTS | H010102012117 (ECS-271.2-18-33-JGN-TR3)
    res = analyze_datasheet(b"", text="SMD 4-SMD, 27.12MHz, ±20ppm, 18pF, 40 Ohms, DTSİS (ECS Inc.)", part_number="H010102012117 (ECS-271.2-18-33-JGN-TR3)", rulebook=rulebook, serial_counter="105")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYSB8105-DTS"

    # 60. CRYTA0003-RA | R38-32.768-12.5-EXT-5PPM
    res = analyze_datasheet(b"", text="TH 32.768kHz, Raltron", part_number="R38-32.768-12.5-EXT-5PPM", rulebook=rulebook, serial_counter="003")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYTA0003-RA"

    # 61. CRYTA0021-ZZ8 | CFS-20632768DZFB
    res = analyze_datasheet(b"", text="TH 32.768kHz,  Citizen", part_number="CFS-20632768DZFB", rulebook=rulebook, serial_counter="021")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYTA0021-ZZ8"

    # 62. CRYTA0046-ISP | 150-0220-010 (AB38T-32.768KHZ)
    res = analyze_datasheet(b"", text="TH 32.768kHz, 12.5pF, 30 kOhms, Cylindrical, ISIPARK (Abracon)", part_number="150-0220-010 (AB38T-32.768KHZ)", rulebook=rulebook, serial_counter="046")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYTA0046-ISP"

    # 63. CRYTA6002-MT | CSTLS4M00G53-A0/B0
    res = analyze_datasheet(b"", text="TH 4MHz, Murata", part_number="CSTLS4M00G53-A0/B0", rulebook=rulebook, serial_counter="002")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYTA6002-MT"

    # 64. CRYTA6077-MT | CSTLS4M00G53-B0
    res = analyze_datasheet(b"", text="TH Resonator, 4MHz, 15 pF ±0.2% 30 Ohms, Murata", part_number="CSTLS4M00G53-B0", rulebook=rulebook, serial_counter="077")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYTA6077-MT"

    # 65. CRYTA7005-CY | 13.560M DIP(HC49S)
    res = analyze_datasheet(b"", text="TH 13.560MHz, CY", part_number="13.560M DIP(HC49S)", rulebook=rulebook, serial_counter="005")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYTA7005-CY"

    # 66. CRYTA8006-CY | 6,7458M DIP(HC49S)
    res = analyze_datasheet(b"", text="TH 6.7458MHz, CY", part_number="6,7458M DIP(HC49S)", rulebook=rulebook, serial_counter="006")
    assert res["material_type"] == "CRY"
    assert res["erp_code"] == "CRYTA8006-CY"

