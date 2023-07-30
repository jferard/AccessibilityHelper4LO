import logging
from typing import List, Optional, cast, Callable, Iterable


class Node:
    _logger = logging.getLogger()

    def __init__(self, value: str, action: Optional[Callable[[], None]],
                 children: List["Node"], parent: Optional["Node"], level: int,
                 previous_sibling: Optional["Node"],
                 next_sibling: Optional["Node"]):
        self.value = value
        self.action = action
        self.children = children
        self.parent = parent
        self.level = level
        self.previous_sibling = previous_sibling
        self.next_sibling = next_sibling

    def execute(self):
        if self.action is not None:
            self._logger.debug("Execute action")
            self.action()

    def previous(self) -> Optional["Node"]:
        n = self.previous_sibling
        if n is None:
            if self.parent is None:
                return None
            else:
                return self.parent
        return n

    def next(self) -> Optional["Node"]:
        n = self
        while n is not None:
            if n.next_sibling:
                return n.next_sibling
            else:
                n = n.parent
        return None

    @staticmethod
    def _short_repr(node: "Node") -> str:
        if node is None:
            return "None"
        else:
            return "{}-{}".format(id(node), node.value)

    def __repr__(self) -> str:
        return ("Node(id={}, value={}, action={}, "
                "children={}, parent={}, level={}, previous_sibling={},"
                "next_sibling={})").format(
            id(self), self.value, self.action,
            [Node._short_repr(c) for c in self.children],
            Node._short_repr(self.parent), self.level,
            Node._short_repr(self.previous_sibling),
            Node._short_repr(self.next_sibling))


class NodeBuilder:
    def __init__(self, value: str,
                 action: Optional[Callable[[], None]] = None):
        self.value = value
        self.action = action
        self.children = cast(List[NodeBuilder], [])
        self.parent = cast(Optional[NodeBuilder], None)
        self.level = -1
        self.previous_sibling = cast(Optional[NodeBuilder], None)
        self.next_sibling = cast(Optional[NodeBuilder], None)

    def append_child(self, node: "NodeBuilder"):
        self.children.append(node)

    def extend_children(self, nodes: Iterable["NodeBuilder"]):
        self.children.extend(nodes)

    def freeze_as_root(self):
        self.level = 0
        self._freeze()

    def _freeze(self):
        self.__class__ = Node
        for i, c in enumerate(self.children):
            c.parent = self
            c.level = self.level + 1
            if i > 0:
                c.previous_sibling = self.children[i - 1]
            else:
                c.previous_sibling = None
            if i < len(self.children) - 1:
                c.next_sibling = self.children[i + 1]
            else:
                c.next_sibling = None
            c._freeze()

    def execute(self):
        raise ValueError()

    @staticmethod
    def _short_repr(node: "NodeBuilder") -> str:
        if node is None:
            return "None"
        else:
            return "{}-{}".format(id(node), node.value)

    def __repr__(self) -> str:
        return "NodeBuilder(id={}, value={}, action={}, children={})".format(
            id(self), self.value, self.action,
            [NodeBuilder._short_repr(c) for c in self.children])


class Tree:
    _logger = logging.getLogger(__name__)

    def __init__(self, root: Node):
        self.root = root
        self.focus = root

    def down(self):
        sibling = self.focus.next_sibling
        if sibling:
            self.focus = sibling

    def up(self):
        sibling = self.focus.previous_sibling
        if sibling:
            self.focus = sibling

    def right(self):
        parent = self.focus.parent
        if parent:
            self.focus = parent

    def left(self):
        children = self.focus.children
        if children:
            self.focus = children[0]

    def enter(self):
        self.focus.execute()

    def text(self, node: Node) -> str:
        s = "+"
        cur = self.focus.parent
        while cur is not None:
            if cur == node:
                s = "-"
            cur = cur.parent
        space_count = (8 + node.level - self.focus.level) * 4
        value = space_count * " " + str(
            node.value)
        if node.children:
            value += " " + s
        return value
