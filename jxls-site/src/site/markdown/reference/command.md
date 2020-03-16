Command
========

Introduction
------------
Command defines a transformation applied to the associated XLS Area.

Jxls uses the following *Command* interface to represent the command.

    public interface Command {
        String getName(); // command name
        List<Area> getAreaList(); // list of command areas
        Command addArea(Area area); // adds an Area to the command
        Size applyAt(CellRef cellRef, Context context); // applies the command at given cell and context
        void reset(); // resets command data for repeatable command usage
        // the next two commands operate on a cell shift mode for the command
        // shift mode currently can take 2 values Command.INNER_SHIFT_MODE (default) and Command.ADJACENT_SHIFT_MODE
        void setShiftMode(String mode);
        String getShiftMode();
    }

Every command has a `name` and `Area` list.

The `name` defines the command name which is used to refer to the command in Excel or XML markup.

The `Area` list is a list of command parameters of  `Area` type.

Built-in commands
-----------------
There are 3 built-in commands in Jxls

* [Each-Command](each_command.html) - for object collection iteration

* [If-Command](if_command.html) - for conditional output

* [Image-Command](image_command.html) - for image output

Custom Commands
---------------
It is easy to create your own commands.

An example of creating a custom command can be found in [Custom command example](../samples/custom_command.html)
