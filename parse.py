import struct


def get_techbase(tb):
    ret = []
    txt_tb = str(tb)
    if '0' in txt_tb:
        ret.append('Common')
    if '1' in txt_tb:
        ret.append('New Republic')
    if '2' in txt_tb:
        ret.append('Imperial')
    if '3' in txt_tb:
        ret.append('Herald')
    if '4' in txt_tb:
        ret.append('Ploxus')
    if not ret:
        raise RuntimeError
    return ret


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
    #with open('engines.db', 'rb') as f:
    #    for (e1, e2, e3, e4, e5, e6, e7, e8) in struct.iter_unpack('<25s7h', f.read()):
    #        print(e1, e2, e3, e4, e5, e6, e7, e8)
        #record = f.read(39)
        #while record:
        #    (e1, e2, e3, e4, e5, e6, e7, e8) = struct.unpack('<25s7h', record)
        #    print(e1, e2, e3, e4, e5, e6, e7, e8)
        #    record = f.read(39)
    #with open('weapons.db', 'rb') as f2:
        #for (w1, w2, w3, w4, w5, w6, w7, w8, w9, w10, w11) in struct.iter_unpack('<25s6sdh15s6s5h', f2.read()):
    #    for record in struct.iter_unpack('<25s6sdh15s6s5h', f2.read()):
    #        (w1, w2, w3, w4, w5, w6, w7, w8, w9, w10, w11) = record
    #        name = record[0].decode()
    #        print(name, w2, w3, w4, w5, w6, w7, w8, w9, w10, w11)
    for record in parse_weapons('weapons.db'):
        print(record)
