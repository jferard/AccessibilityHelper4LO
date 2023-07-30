import unittest

from ah4lo_tree import NodeBuilder


class NodeTestCase(unittest.TestCase):
    def test_node(self):
        """
        A -- B ---- C
          |  \----- D
          \- E ---- F
             \----- G
             \----- H
        """
        a = NodeBuilder("A")
        b = NodeBuilder("B")
        c = NodeBuilder("C")
        d = NodeBuilder("D")
        e = NodeBuilder("E")
        f = NodeBuilder("F")
        g = NodeBuilder("G")
        h = NodeBuilder("H")
        a.append_child(b)
        a.append_child(e)
        b.append_child(c)
        b.append_child(d)
        e.append_child(f)
        e.append_child(g)
        e.append_child(h)

        print(a)
        a.freeze_as_root()
        print(a)

        print(b)
        print(e)

        print(c)
        print(d)

        print(f)
        print(g)
        print(h)


if __name__ == '__main__':
    unittest.main()
