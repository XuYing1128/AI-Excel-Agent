from excel_agent import formula_library as F


def test_common_formulas():
    assert F.excel_sum("A1:A10") == "=SUM(A1:A10)"
    assert F.average("B1:B3") == "=AVERAGE(B1:B3)"
    assert F.countifs(("A:A", "饮品"), ("B:B", ">0")) == '=COUNTIFS(A:A,"饮品",B:B,">0")'
    assert F.sumifs("C:C", ("A:A", "华东")) == '=SUMIFS(C:C,A:A,"华东")'
    assert F.gross_margin_from_profit("I4", "H4") == "=IFERROR(I4/H4,0)"
    assert F.growth_rate("B5", "B4") == "=IFERROR((B5-B4)/B4,0)"
    assert F.roi("F4", "C4+D4") == "=IFERROR((F4-(C4+D4))/(C4+D4),0)"


def test_guarded_formula():
    assert F.guarded(["A4"], "=B4*C4") == '=IF(COUNTA(A4)=0,"",B4*C4)'
