"""Convert a starfighter data file in VB6 format to JSON."""

import itertools
import json
import os
import sys
from vb6_stuff import *


def new_starfighter():
    return {
        'name': '',
        'abbr': '',
        'space': 0,
        'criticals': {},
        'armor': {},
        'shields': 0,
        'speed': 0,
        'techbase': '',
        'wings': 0
    }


def new_criticals():
    return {
        'C': [],
        'F': [],
        'LW': [],
        'RW': []
    }


def get_armor_type(arm):
    if arm == 0:
        return 'Standard'
    elif arm == 1:
        return 'Didrate'
    elif arm == 2:
        return 'Trinnium'
    elif arm == 3:
        return 'Tri-Di Composite'
    elif arm == 4:
        return 'Clearplast'
    else:
        raise RuntimeError


def location_name(loc):
    if loc == 1:
        return 'C'
    elif loc == 2:
        return 'F'
    elif loc == 3:
        return 'LW'
    elif loc == 4:
        return 'RW'


# This is taken from itertools documentation.
def grouper(iterable, n, fillvalue=None):
    """Collect data into fixed-length chunks or blocks"""
    # grouper('ABCDEFG', 3, 'x') --> ABC DEF Gxx"
    args = [iter(iterable)] * n
    return itertools.zip_longest(*args, fillvalue=fillvalue)


def parse_criticals(crit_buf, weapons, engines):
    crits = new_criticals()
    for (loc, rec_num, id_num) in grouper(crit_buf, 3):
        loc_name = location_name(loc)
        if not loc_name:
            continue
        if rec_num in weapons:
            name = 'W' + str(rec_num) + '-' + weapons[rec_num]
            crits[loc_name].append(name)
        elif rec_num < 0 and -rec_num in engines:
            # engine criticals are stored with negative numbers
            name = 'E' + str(-rec_num) + '-' + engines[-rec_num]
            crits[loc_name].append(name)
        elif rec_num == 0 and loc == 1:
            # this looks like a bug that a blank critical was added to the cockpit
            continue
        else:
            raise RuntimeError
    return crits


def parse_starfighter(filename, weapons, engines):
    with open(filename, 'rb') as f:
        record = read_unpack(f, '<25s6s94h')
        return {
            'name': record[0].decode().strip(),
            'abbr': record[1].decode().strip(),
            'space': (record[2] + 2) * 5,
            'criticals': parse_criticals(record[3:86], weapons, engines),
            'armor': {
                'C': record[87],
                'F': record[88],
                'LW': record[89],
                'RW': record[90],
                'type': get_armor_type(record[91])
            },
            'shields': record[92],
            'speed': record[93],
            'techbase': get_techbase(record[94])[0],
            'wings': record[95]
        }


def parse_old_criticals(sws_file):
    crits = new_criticals()
    locs = ['C', 'F', 'LW', 'RW']
    items_seen = {}
    for loc, _ in zip(itertools.cycle(locs), range(48)):
        crit_record = read_unpack(sws_file, '<21shf')
        num_criticals = crit_record[1]
        if num_criticals == 0:
            continue
        item = crit_record[0].decode().strip()
        name_to_use = item
        if item in items_seen:
            items_seen[item] += 1
            name_to_use = item + ' #' + str(items_seen[item])
        else:
            items_seen[item] = 1
        crits[loc].extend([name_to_use] * num_criticals)
    return crits


def parse_old_starfighter(filename, weapons, engines):
    sf = new_starfighter()
    with open(filename, 'rb') as f:
        engine_record = read_unpack(f, '<h20s2hfh')
        sf['speed'] = engine_record[2]

        armor_record = read_unpack(f, '<16s5hf')
        sf['armor'] = {
            'C': armor_record[1],
            'F': armor_record[2],
            'LW': armor_record[3],
            'RW': armor_record[4],
            'type': armor_record[0].decode().strip()
        }

        ship_record = read_unpack(f, '<25s6s3hc4h')
        sf['name'] = ship_record[0].decode().strip()
        sf['abbr'] = ship_record[1].decode().strip()
        sf['space'] = ship_record[2]
        sf['shields'] = ship_record[3]
        sf['techbase'] = get_techbase(ship_record[5])[0]
        sf['wings'] = ship_record[6]

        # Skip the internal structure record.
        read_unpack(f, '<f4h')

        sf['criticals'] = parse_old_criticals(f)
    return sf


def load_data_file(filename):
    with open(filename, 'r') as f:
        for obj in json.load(f):
            yield (obj['id'], obj['name'])


def script_path():
    return os.path.dirname(os.path.abspath(__file__))


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('usage:', 'python', sys.argv[0], '<sw2 file>')
        sys.exit(1)

    mydir = script_path()
    weapons_file = os.path.join(mydir, 'weapons.json')
    weapons = dict(load_data_file(weapons_file))
    engines_file = os.path.join(mydir, 'engines.json')
    engines = dict(load_data_file(engines_file))

    _, ext = os.path.splitext(sys.argv[1])
    if ext == '.sw2':
        starfighter = parse_starfighter(sys.argv[1], weapons, engines)
        print(json.dumps(starfighter, indent=4))
    else:
        starfighter = parse_old_starfighter(sys.argv[1], weapons, engines)
        print(json.dumps(starfighter, indent=4))
