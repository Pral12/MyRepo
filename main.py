def check_exist_attrs(obj, *args):
    return {i: i in dir(obj) for i in list(args)[0]}


def create_attrs(obj, *args):
    print(dict(list(args)[0]))
    for key, value in dict(list(args)[0]).items():
        setattr(obj, key, value)


def print_goods(lst):
    pass

create_attrs(print_goods, [('house', 1), ('level', 3), ('cost', 1000000)])
print(check_exist_attrs(print_goods, ['house', 'level', 'cost']))
print(print_goods.house)
print(print_goods.level)
print(print_goods.cost)
