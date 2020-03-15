Expression Language
===================

Overview
--------
By default Jxls uses [Apache JEXL](http://commons.apache.org/proper/commons-jexl/) expression language to evaluate property expressions specified in Excel template file.

See [JEXL Syntax Reference](http://commons.apache.org/proper/commons-jexl/reference/syntax.html) to see what expressions can be used.

Customize Jexl processing
-----------------------------
If you need to customize Jexl processing you can get a reference to [JexlEngine](https://commons.apache.org/proper/commons-jexl/javadocs/apidocs-2.1/org/apache/commons/jexl2/JexlEngine.html) from *Transformer* and apply necessary configuration.

For example the following code registers a custom Jexl function under `demo` namespace
 
         Transformer transformer = TransformerFactory.createTransformer(is, os);
         ...
         JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
         Map<String, Object> functionMap = new HashMap<>();
         functionMap.put("demo", new JexlCustomFunctionDemo());
         evaluator.getJexlEngine().setFunctions(functionMap);

Here *JexlCustomFunctionDemo* class has a method

    public Integer mySum(Integer x, Integer y){
        return x + y;
    }
    
So in your template you can use this function like
     
     ${demo:mySum(x,y)}
     
where `x` and `y` are parameters from *Context*
     
See JexlCustomFunctionDemo.java for a full example.      


Change Expression Engine
--------------------------
 
You may prefer not to use [Apache JEXL](http://commons.apache.org/proper/commons-jexl/) but use some other expression processing engine 
(for example use [Spring Expression Language (SpEL)](http://docs.spring.io/spring/docs/current/spring-framework-reference/html/expressions.html) ).

Jxls allows you to substitute the default evaluation engine with the one you prefer.

To do this you should just implement a single method of *ExpressionEvaluator* interface to delegate the expression evaluation processing to the engine you want 

    public interface ExpressionEvaluator {
        Object evaluate(String expression, Map<String,Object> context);
    }
        
And then you need to pass your *ExpressionEvaluator* implementation to *TransformationConfig* as shown in the below code
          
      ExpressionEvaluator evaluator = new MyCustomEvaluator(); // your own implementation based for example on SpEL
      transformer.getTransformationConfig().setExpressionEvaluator(evaluator);    
      
      