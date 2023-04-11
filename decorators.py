import time

# performance decorator
def timeit(func):
    def wrapper(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()
        print(f"Time taken to execute {func.__name__} is {round(end - start, 4)} seconds")
        return result
    return wrapper