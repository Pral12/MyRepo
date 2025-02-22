from functools import wraps

def add_attrs(**kwags):
    def decorator(func):

        @wraps(func)
        def inner(*args, **kwargs):
            return func(*args, **kwargs)

        for k, v in kwags.items():
            setattr(inner, k, v)
        return inner

    return decorator


@add_attrs(test=True, ordered=True)
def add(a, b):
    return a + b

print(add(10, 5))
print(add.test)
print(add.ordered)


@add_attrs(hello='World', marks=[1, 2, 3], cash=100)
def add(a, b):
    return a + b

print(add(10, 5))
print(add.hello)
print(add.marks)
print(add.cash)
print(getattr(add, 'ordered', None))

