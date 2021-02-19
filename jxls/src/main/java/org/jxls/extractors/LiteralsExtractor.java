package org.jxls.extractors;

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
    
    public List<String> extract(final String text) {
        List<String> literalList = new ArrayList<>();
        String literal = "";
        for (int i = 0, n = text.length(); i < n; i++) {
            if (literal.endsWith(XlsCommentAreaBuilder.COMMAND_PREFIX)) {
                
                // if the first 3 chars of literal are 'jx:'
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
                    .filter(line -> !line.isEmpty())
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
                    if (!quotationsStack.empty() && quotationsStack.peek() == symbol) {
                        quotationsStack.pop();
                    } else {
                        quotationsStack.push(symbol);
                    }
                    break;
                case '(':
                    if (quotationsStack.empty()) {
                        bracketsStack.push(symbol);
                    }
                    break;
                case ')':
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
                    break;
            }

            cmd += symbol;
        }
        return cmd;
    }
}
