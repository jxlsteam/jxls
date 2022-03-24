package org.jxls.util;

import java.util.Stack;

public class SymbolManager {
    public SymbolManager() {
    }

    String getString(String text, String cmd, Stack<Character> bracketsStack, Stack<Character> quotationsStack, int i, char symbol) {
        if (quotationsStack.empty()) {
            if (bracketsStack.peek() == '(') {
                bracketsStack.pop();
                if (bracketsStack.empty()) {

                    return cmd + symbol;
                }
            } else {
                // no opening, but closing !!
                String errorMessage = "Got closing ) but no opening of it. Pos: " + text.substring(0, i);
                throw new IllegalArgumentException(errorMessage);
            }
        }
        return null;
    }

    void extracted(Stack<Character> bracketsStack, Stack<Character> quotationsStack, char symbol) {
        if (quotationsStack.empty()) {
            bracketsStack.push(symbol);
        }
    }

    void extracted(Stack<Character> quotationsStack, char symbol) {
        if (!quotationsStack.empty() && quotationsStack.peek() == symbol) {
            quotationsStack.pop();
        } else {
            quotationsStack.push(symbol);
        }
    }
}