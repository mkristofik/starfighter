"""Helper functions for reading the old VB6 data files."""


def get_techbase(tb):
    ret = []
    txt_tb = str(tb)
    if '0' in txt_tb:
        ret.append('Common')
    if '1' in txt_tb or 'N' in txt_tb:
        ret.append('New Republic')
    if '2' in txt_tb or 'I' in txt_tb:
        ret.append('Imperial')
    if '3' in txt_tb or 'H' in txt_tb:
        ret.append('Herald')
    if '4' in txt_tb or 'P' in txt_tb:
        ret.append('Ploxus')
    if not ret:
        raise RuntimeError
    return ret
