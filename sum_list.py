def get_total(lst: list[int | list[int]]) -> int:
    if len(lst) == 0:
        return 0
    if isinstance(lst[0], int):
        return get_total(lst[1::]) + lst[0]
    elif isinstance(lst[0], list):
        return get_total(lst[0]) + get_total(lst[1::])




print(get_total([[1, 2, 3], [4, 5], [6, 7, 8]]))

print(get_total([1, 2, 3, 4, 5]))
print(get_total([1, [2, [3, [4, 5]]]]))

