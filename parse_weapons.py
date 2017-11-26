"""Convert a weapons data file in VB6 format to JSON."""

import json
import struct
import sys
from vb6_stuff import *


def get_locations(locs):
    ret = []
    txt_locs = str(locs)
    if '1' in txt_locs:
        ret.append('cockpit')
    if '2' in txt_locs:
        ret.append('fuselage')
    if '3' in txt_locs:
        ret.append('left wing')
    if '4' in txt_locs:
        ret.append('right wing')
    if not ret:
        raise RuntimeError
    return ret


def get_options(opts):
    if opts == 0:
        return []
    elif opts == 2:
        return ['weapon']
    elif opts == 12:
        return ['warhead launcher', 'weapon']
    else:
        raise RuntimeError


def parse_weapons(filename):
    with open(filename, 'rb') as f:
        count = 0
        for record in struct.iter_unpack('<25s6sdh15s6s5h', f.read()):
            count += 1
            is_deleted = record[10]
            if is_deleted:
                continue
            yield {'id': count,
                'name': record[0].decode().strip(),
                'damage': record[1].decode().strip(),
                'space': record[2],
                'criticals': record[3],
                'range': record[4].decode().strip(),
                'tohit': record[5].decode().strip(),
                'maxnum': record[6],
                'techbase': get_techbase(record[7]),
                'locations': get_locations(record[8]),
                'options': get_options(record[9])}


if __name__ == '__main__':
    filename = 'weapons.db'
    if len(sys.argv) > 1:
        filename = sys.argv[1]

    print(json.dumps(list(parse_weapons(filename)), indent=4))
