import sys
sys.path.append("z:/PCP/PROJETOS MARLON/ProgramarProd")
import data_manager as dm

def run_tests():
    test_cases = [
        # (bloco_a, bloco_b, expected_match)
        ("4244", "4244/771418", True),
        ("771418", "4244/771418", True),
        ("4244/771418", "4244", True),
        ("4244/771418", "771418", True),
        ("4244/771418", "4244/771418", True),
        ("4244.0", "4244/771418", True),
        ("4244/771418.0", "4244", True),
        ("4245", "4244", False),
        ("4245", "4244/771418", False),
    ]

    success = True
    for a, b, exp in test_cases:
        res = dm.blocos_match(a, b)
        if res != exp:
            print(f"FAIL: blocos_match({repr(a)}, {repr(b)}) -> expected {exp}, got {res}")
            success = False
        else:
            print(f"PASS: blocos_match({repr(a)}, {repr(b)}) -> {res}")
            
    if success:
        print("\nAll unit tests passed successfully!")
    else:
        print("\nSome unit tests failed!")

if __name__ == "__main__":
    run_tests()
