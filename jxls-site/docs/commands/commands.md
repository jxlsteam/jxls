---
sidebar_position: 1
---

# Commands

## Notes and command syntax

Commands are written as Excel notes. You create a note by right-clicking > New Note. A note usually contains
the author on the first line. This is followed by commands line by line, with commands at the beginning
of the line and the command name beginning with `jx:`.

The command format is like this:

```
"jx:" + command name + "(" + attribute name + "=\"" + value + "\" " [...] + ")\n"
```

example:

```
jx:each(items="employees" var="e" lastCell="C2")
```

Some write it like this, which is also fine:

```
jx:each(items="employees", var="e", lastCell="C2")
```

## lastCell

Each command refers to a cell area. The cell area starts at the cell that has the note (top left corner) and
ends at the cell specified in the lastCell attribute (bottom right corner).

## Commands included

Jxls supports these built-in and ready-to-use commands:

- jx:area
- [jx:each](each)
- [jx:if](if)
- [jx:grid](grid)
- [jx:updateCell](update-cell)
- [jx:params](params)
- [jx:image](image)
- [jx:mergeCells](merge-cells)

`jx:area` defines the cell area that Jxls should process. jx:area is therefore usually found in cell A1 and
its lastCell attribute defines the lower right corner of the worksheet area used.

`jx:each` is primarily for creating rows. It's the most important command. jx:each can also create columns or sheets.

`jx:if` is for showing and hiding cell areas based on a condition.

`jx:grid` creates dynamically a grid out of column headers and data objects.

`jx:updateCell` can be used for applying individual processing instructions for modifying a cell area.

`jx:params` is a special command for setting a parameter.

`jx:image` is for adding an image to the sheet.

`jx:mergeCells` is for combining cells to one new cell.
