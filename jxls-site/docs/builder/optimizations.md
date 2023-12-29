---
sidebar_position: 3
---

# Optimizations

Sometimes reporting takes too long. What optimization options does Jxls offer in terms of speed and memory requirements?

## Streaming

In production streaming is the most important option for an optimization. Usually you don't need other options.

See [streaming](streaming)

The sheet order can have an impact on the speed. Place sheets with streaming and simple or less formulas at the begin.

## FastFormulaProcessor

See [builder options](../builder).

## Recalculate formulas

See [builder options](../builder).

## JexlExpressionEvaluator

Use the thread-local JexlExpressionEvaluator because it caches engine and expressions. That makes reporting faster.
JexlExpressionEvaluator is the default. See [expressions](../expressions).

## Lean memory management

If you put an list in the Jxls data map which used for iteration and holds fat objects you can build yourself some
logic to free those fat objects after one object is used. This only works if the list is used only once in the template.
