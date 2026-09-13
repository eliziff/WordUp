"""Defect controls for the development comparison oracle; no Word required."""
import copy
from types import SimpleNamespace as Object
import unittest

from compare import form_snapshot, require_form_snapshot


class FormSnapshotTest(unittest.TestCase):
    def test_mismatch_names_the_first_structural_identity(self):
        expected = {("Form1", "Button1", (), 0): ("before",)}
        actual = {("Form1", "Button1", (), 0): ("after",)}
        with self.assertRaisesRegex(AssertionError, r"Form1.*Button1.*expected.*before.*got.*after"):
            require_form_snapshot(actual, expected, "candidate")

    def test_unnamed_nested_records_keep_distinct_structural_identities(self):
        def unnamed():
            return Object(name="", kind="MSForms.Page", id=0, tab_index=None,
                          properties=lambda: {}, children=(), record=None)
        first, second = unnamed(), unnamed()
        form = Object(name="Form1", controls=(first, second), properties=lambda: {},
                      designer_source="original", _levels=[])
        snapshot = form_snapshot(Object(forms=lambda: [form]))
        self.assertEqual(len(snapshot), 3)
        self.assertIn(("Form1", "", (), 0), snapshot)
        self.assertIn(("Form1", "", (), 1), snapshot)

    def test_detects_reparenting_reordering_and_same_length_picture_changes(self):
        def control(name, children=()):
            return Object(name=name, kind="MSForms.Label", id=name, tab_index=0,
                          properties=lambda: {"Caption": "unchanged"}, children=children,
                          record=Object(pictures={"Picture": b"abcd"}))

        child = control("child")
        parent = control("parent", (child,))
        form = Object(name="Form1", controls=(parent,), properties=lambda: {},
                      designer_source="original", _levels=[])
        document = Object(forms=lambda: [form])
        expected = form_snapshot(document)
        self.assertEqual(expected, form_snapshot(document))
        parent.children = ()
        form.controls = (parent, child)
        self.assertNotEqual(expected, form_snapshot(document))
        ordered = form_snapshot(document)
        form.controls = (child, parent)
        self.assertNotEqual(ordered, form_snapshot(document))
        form.controls = (parent,)
        parent.children = (child,)
        child.record.pictures["Picture"] = b"abce"
        self.assertNotEqual(expected, form_snapshot(document))
        child.record.pictures["Picture"] = b"abcd"
        self.assertEqual(expected, form_snapshot(document))
        child.properties = lambda: {"Caption": "changed"}
        self.assertNotEqual(expected, form_snapshot(document))
        adapted = copy.deepcopy(expected)
        adapted["Form1", "child", (0,), 0][3]["Caption"] = "changed"
        self.assertEqual(adapted, form_snapshot(document))


if __name__ == "__main__":
    unittest.main()
