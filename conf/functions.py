def unpack_single_item_lists(val):
    if isinstance(val, list) and len(val) == 1:
        return val[0]
    else:
        return val
