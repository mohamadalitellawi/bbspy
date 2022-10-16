from pathlib import Path
import csv

def main():
    p = Path('.')
    print(f"Current Directory is\t{p.resolve()}")

    # Listing subdirectories:
    [print(x) for x in p.iterdir() if x.is_dir()]

    # Listing Python source files in this directory tree:
    pp = list(p.glob('**/*.py'))
    print(pp)

    # Navigating inside a directory tree:
    check_path = p / "wip" / "IMG"
    print(check_path.exists())

    print(check_path.name)
    print(check_path.resolve().parts)
    print(check_path.resolve().parent)
    print(check_path.resolve().parents[1])

    file = check_path / 'BAR-01.WMF'
    print(file.name)
    print(file.stem)
    print(file.suffix)


    print(Path.cwd())
    print(Path.home())

    # print(sorted(Path('.').glob('*.py')))
    # print(sorted(Path('.').glob('*/*.py')))
    # print(sorted(Path('.').glob('**/*.py')))

    pp = Path.home()
    for child in pp.iterdir(): print(child)

    pp = Path('./wip/PARAMETER NUMBERING.csv')
    with pp.open() as f:
        print(f.readline())

    f = Path('tmp.txt')
    f.write_text("Hello World!")
    print(f.read_text())


def get_parameters(parameter_file:Path):
    with open(parameter_file.resolve(), newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        parameters = []
        for row in reader:
            parameters.append(row)
        return parameters




if __name__ == "__main__":
    main()