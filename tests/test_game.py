import json
import os
import tempfile
import random
import unittest

from quien_millonario import load_questions, get_5050_indices


class TestGameLogic(unittest.TestCase):
    def test_load_questions_contains_entries(self):
        # Use a temp path to avoid mutating real data
        with tempfile.TemporaryDirectory() as td:
            # copy data file to temp
            src = os.path.join(os.path.dirname(__file__), '..', 'data', 'questions.json')
            dst = os.path.join(td, 'questions.json')
            with open(src, 'r', encoding='utf-8') as f:
                data = f.read()
            with open(dst, 'w', encoding='utf-8') as f:
                f.write(data)
            qs = load_questions(dst)
            self.assertTrue(len(qs) > 0)

    def test_5050_includes_correct(self):
        with open(os.path.join(os.path.dirname(__file__), '..', 'data', 'questions.json'), 'r', encoding='utf-8') as f:
            qlist = json.load(f)
        q = qlist[0]
        rng = random.Random(42)
        a, b = get_5050_indices(q, rng)
        self.assertIn(a, [0,1,2,3])
        self.assertIn(b, [0,1,2,3])
        self.assertIn(q['answer'], [a, b])  # one of them must be correct


if __name__ == '__main__':
    unittest.main()
