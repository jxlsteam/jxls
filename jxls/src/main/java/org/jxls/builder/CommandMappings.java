package org.jxls.builder;

import org.jxls.command.Command;

public interface CommandMappings {
    
    void addCommandMapping(String commandName, Class<? extends Command> commandClass);

    void removeCommandMapping(String commandName);
    
    Class<? extends Command> getCommandClass(String commandName);
}
