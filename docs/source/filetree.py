from pathlib import Path

# prefix components:
space = '    '
branch = '│   '
# pointers:
tee = '├── '
last = '└── '
first = '── '


def get_tree_string(directory):

    p = Path(directory)

    tree_string = first + p.name + "\n"
    prefix = '    '  
    for line in tree(p, prefix):
        tree_string += line + "\n"

    return tree_string


def tree(dir_path: Path, prefix: str = ''):
    """A recursive generator, given a directory Path object
    will yield a visual tree structure line by line
    with each line prefixed by the same characters

    Source: https://stackoverflow.com/a/59109706
    """
    contents = list(dir_path.iterdir())
    # contents each get pointers that are ├── with a final └── :
    pointers = [tee] * (len(contents) - 1) + [last]
    for pointer, path in zip(pointers, contents):
        yield prefix + pointer + path.name
        if path.is_dir():  # extend the prefix and recurse:
            extension = branch if pointer == tee else space
            # i.e. space because last, └── , above so no more |
            yield from tree(path, prefix=prefix + extension)
