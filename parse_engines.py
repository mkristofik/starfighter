"""Convert an engines data file in VB6 format to JSON."""

import json
import struct
from vb6_stuff import *


def get_engine_type(typenum):
    if typenum == 0:
        return 'standard engine'
    elif typenum == 1:
        return 'goofy engine'
    elif typenum == 2:
        return 'hyperdrive'
    elif typenum == 3:
        return 'afterburner'
    else:
        raise RuntimeError


def parse_engines():
    with open('engines.db', 'rb') as f:
        count = 0
        for record in struct.iter_unpack('<25s7h', f.read()):
            count += 1
            is_deleted = record[7]
            if is_deleted:
                continue
            yield {'id': count,
                'name': record[0].decode().strip(),
                'criticals': record[1],
                'type': get_engine_type(record[2]),
                'techbase': get_techbase(record[3]),
                'base_maneuverability': record[4],
                'rating_speed_modifier': record[5],
                'speed_mult_pct': record[6]}


if __name__ == '__main__':
    print(json.dumps(list(parse_engines()), indent=4))
