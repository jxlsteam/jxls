package org.jxls.extractors;

import static java.util.stream.Collectors.toList;

import java.util.Arrays;
import java.util.List;
import java.util.Stack;

import org.jxls.builder.xls.XlsCommentAreaBuilder;

public class LiteralsExtractor {
	public static final int COMMAND_PREFIX_LENGTH = 3;
    private String text;
    private List<String> literalList;
	
    public LiteralsExtractor(String text, List<String> literalList) {
    	this.literalList = literalList;
    	this.text = text;
    }
    
    public List<String> extract() {
    	String literal = "";
		for (int i = 0, n = text.length(); i < n; i++) {

			// if the first 3 chars of literal are 'jx:'
			if (literal.length() >= COMMAND_PREFIX_LENGTH && literal.endsWith(XlsCommentAreaBuilder.COMMAND_PREFIX)) {
				// save comment parsed before
				literal = literal.substring(0, literal.length() - COMMAND_PREFIX_LENGTH);
				makeCommentLines(literal, literalList);
				String remainingText = text.substring(i - COMMAND_PREFIX_LENGTH, n);
				String cmd = makeOperators(remainingText);
				literal = "";
				literalList.add(cmd);
				i = i + cmd.length() - COMMAND_PREFIX_LENGTH;

				// if text no ending
				if (i != n) {
					literal = String.valueOf(text.charAt(i));					
				}
				else {
					literal = "";					
				}

			} else {
				// else is the next char to new literal
				literal = literal + String.valueOf(text.charAt(i));
			}
		}

		// last comment
		if (!literal.isEmpty()) {
			makeCommentLines(literal, literalList);
		}
		return literalList;
    }
    
    private void makeCommentLines(String literal, List<String> literalList) {
		literalList.addAll(
				Arrays.stream(literal.split("\n"))
					.filter(line -> ! line.isEmpty())
					.collect(toList()));
	}
	
	private String makeOperators(String text) {
		// result command
		String cmd = "";
		Stack<Character> bracketsStack = new Stack<>();
		Stack<Character> quotationsStack = new Stack<>();

		for (int i = 0; i < text.length(); i++) {
			char symbol = text.charAt(i);

			if (symbol == '"' || symbol == '\'') {
				if (!quotationsStack.empty() && quotationsStack.peek() == symbol) {
					quotationsStack.pop();					
				}
				else {
					quotationsStack.push(symbol);									
				}
			}

			switch (symbol) {
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
			default:
				break;
			}

			cmd = cmd + symbol;

		}

		return cmd;
	}
    
}
