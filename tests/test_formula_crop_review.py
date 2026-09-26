import sys
from pathlib import Path
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from review_formula_crops import inspect_crop


def main():
    image = Image.new('RGB', (100, 80), 'white')
    ImageDraw.Draw(image).rectangle((25, 25, 60, 50), fill='black')
    _, ink, edges = inspect_crop(image, [20, 20, 50, 40])
    assert ink == [5, 5, 41, 31] and edges == []
    _, ink, edges = inspect_crop(image, [30, 20, 30, 40])
    assert edges == ['left', 'right']
    assert inspect_crop(image, [0, 0, 10, 10])[1] is None
    for box in ([-1, 0, 5, 5], [90, 0, 20, 5], [0, 0, 0, 5], [0, 0, float('nan'), 5]):
        try:
            inspect_crop(image, box)
        except ValueError:
            pass
        else:
            raise AssertionError(f'Invalid crop accepted: {box}')


if __name__ == '__main__':
    main()
