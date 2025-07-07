import database_functions as dbf
import gui_functions as gui_f
import pytest

def test_str2float():
    assert gui_f.str2float('2.5') == 2.5
    assert gui_f.str2float('  -2.5  \n  ') == -2.5
    assert gui_f.str2float('2,5') == 2.5
    assert gui_f.str2float(-3) == -3.0
    with pytest.raises(ValueError): gui_f.str2float('a')
