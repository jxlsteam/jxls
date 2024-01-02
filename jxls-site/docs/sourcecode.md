---
sidebar_position: 6
---

# Source code

Jxls source code is available on Github at https://github.com/jxlsteam/jxls.

## Contributing

1. Create an issue.
2. Fork the Github repository.
3. Clone it into your IDE. A JDK 17 must be configured. Don't use a JRE.
4. Create a feature branch, e.g. feature-47
5. Write a testcase and modify the code.
6. Commit and push the feature branch.
7. The push message should give you a link to execute for creating the pull request.

## Guidelines for creating an issue in Jxls bug tracker

It is very much recommended to provide a working example which demonstrates the issue. Ideally you should raise a pull request to
Jxls repository demonstrating the issue. Providing an example speeds up the identification of the root cause and the resolution of the issue.

If you are unable to provide an example please describe your issue in detail. Provide an excerpt of the code on how you run Jxls
transformation and also attach an input Excel template and an output Excel file you get. Also attach an Excel file with the desired output.

Also in the description please mention which Jxls version you used to reproduce the issue and if you use Microsoft Excel (recommended) or another program.

## Eclipse

1. Clone repository.
2. File > Import > Existing Maven projects
3. Choose parent folder of your Jxls repo folder.
4. Projects will be imported as nested projects.

## Codestyle

- Java 17 syntax (starting from Jxls 3.0.0)
- No wildcards in `import` lines.
- Closing `{` at end of line.
- Use spaces instead of tabs.
- Use @Override annotation.
- Test case class name must end with 'Test'.

## Writing template based testcases

Use **Jxls3Tester** for writing a template based testcase. The template file must have the name of the test class and must be in
the same package in source folder 'src/test/resources'
(e.g. `/jxls-poi/src/test/resources/org/jxls3/EachTest.xlsx`).
The method test() creates the output file in the target folder with _output as part of the name.
Use Jxls3Tester.getWorkbook() for verifying the output file.

```
// Test
Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance());

// Verify
try (TestWorkbook w = tester.getWorkbook()) {
```

If each test method of your testclass use its own template file you can speficy the method name in the xlsx() call as the 2nd argument.

Use xls() if you must use the old .xls format. Prefer the .xlsx format.
