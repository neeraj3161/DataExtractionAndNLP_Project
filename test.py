for x in range(100):
    print(f"Hello World {x} times !!!", end="\r")
    print("\033c", end="", flush=True)