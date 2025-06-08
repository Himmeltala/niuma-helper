def colored(text, color='red', bold=False):
    colors = {
        'black': 30,
        'red': 31,
        'green': 32,
        'yellow': 33,
        'blue': 34,
        'purple': 35,
        'cyan': 36,
        'white': 37
    }
    mode = '1;' if bold else ''
    return f"\033[{mode}{colors[color]}m{text}\033[0m"
