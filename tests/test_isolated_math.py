import sys
from pathlib import Path
import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from measure_isolated_math import compare


def main():
    source = np.full((100, 150, 3), 255, dtype=np.uint8)
    rendered = source.copy()
    source[30:40, 50:70] = 0
    rendered[45:65, 80:120] = 0  # deliberately outside the source crop
    row = compare(source, rendered, [45, 25, 30, 20])
    assert row["ratio"] == 0.5
    assert row["dx_px"] == -40 and row["dy_px"] == -20
    for invalid in ([0, 0, 5, 5], [-1, 0, 30, 20], [100, 80, 100, 40]):
        try:
            compare(source, rendered, invalid)
        except ValueError:
            pass
        else:
            raise AssertionError("Invalid/empty source crop accepted")
    rendered[0:10, 0:10] = 0
    try:
        compare(source, rendered, [45, 25, 30, 20])
    except ValueError:
        pass
    else:
        raise AssertionError("Canvas-clipped ink accepted")


if __name__ == "__main__":
    main()
