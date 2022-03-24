package org.jxls.util;

import static java.util.stream.Collectors.toList;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Stack;

import org.jxls.builder.xls.XlsCommentAreaBuilder;

/**
 * @author Alexander Lust
 */
public class LiteralsExtractor {
    private static final int COMMAND_PREFIX_LENGTH = XlsCommentAreaBuilder.COMMAND_PREFIX.length();
    private final SymbolManager symbolManager = new SymbolManager();

    public List<String> extract(final String text) {
        List<String> literalList = new ArrayList<>();
        String literal = "";
        for (int i = 0, n = text.length(); i < n; i++) {
            String[] lineElems = literal.split("\\n");
            String[] jxLineElem;
            if (lineElems.length == 0) {
                jxLineElem = new String[] {};
            } else {
                jxLineElem = lineElems[lineElems.length - 1].split("jx:");
            }

            if (literal.endsWith(XlsCommentAreaBuilder.COMMAND_PREFIX) && !(jxLineElem.length > 0 && jxLineElem[0].contains("//"))) {
                // if the first 3 chars of literal are 'jx:'
                // is this line commented?
                // save comment parsed before
                literal = literal.substring(0, literal.length() - COMMAND_PREFIX_LENGTH);
                makeCommentLines(literal, literalList);
                String remainingText = text.substring(i - COMMAND_PREFIX_LENGTH, n);
                String cmd = makeOperators(remainingText);
                literal = "";
                literalList.add(cmd);
                i += cmd.length() - COMMAND_PREFIX_LENGTH;

                if (i != n) {
                    // if text no ending
                    literal = String.valueOf(text.charAt(i));
                } else {
                    literal = "";
                }
            } else {
                // else is the next char to new literal
                literal += String.valueOf(text.charAt(i));
            }
        }

        if (!literal.isEmpty()) {
            // last comment
            makeCommentLines(literal, literalList);
        }
        return literalList;
    }
    
    private void makeCommentLines(String literal, List<String> literalList) {
        literalList.addAll(
                Arrays.stream(literal.split("\n"))
                    .filter(line -> !line.replace("\r", "").trim().isEmpty())
                    .collect(toList()));
    }
    
    private String makeOperators(final String text) {
        String cmd = ""; // result command
        Stack<Character> bracketsStack = new Stack<>();
        Stack<Character> quotationsStack = new Stack<>();

        for (int i = 0; i < text.length(); i++) {
            char symbol = text.charAt(i);

            switch (symbol) {
                case '"':
                case '\'':
                    symbolManager.extracted(quotationsStack, symbol);
                    break;
                case '(':
                    symbolManager.extracted(bracketsStack, quotationsStack, symbol);
                    break;
                case ')':
                    String cmd1 = symbolManager.getString(text, cmd, bracketsStack, quotationsStack, i, symbol);
                    if (cmd1 != null) return cmd1;
                    break;
            }

            cmd += symbol;
        }
        return cmd;
    }

    private String getString(String text, String cmd, Stack<Character> bracketsStack, Stack<Character> quotationsStack, int i, char symbol) {
        return symbolManager.getString(text, cmd, bracketsStack, quotationsStack, i, symbol);
    }

    private void extracted(Stack<Character> bracketsStack, Stack<Character> quotationsStack, char symbol) {
        symbolManager.extracted(bracketsStack, quotationsStack, symbol);
    }

    private void extracted(Stack<Character> quotationsStack, char symbol) {
        symbolManager.extracted(quotationsStack, symbol);
    }
}
