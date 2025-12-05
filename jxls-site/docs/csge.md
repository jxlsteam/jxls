# Find missing CellStyle General unit test

If you don't want to auto-fix the missing CellStyle General bug (using `withCellStyleGeneralEnsurer()`),
you could just use a unit test for finding all cells which have missing CellStyle General.
Here is a short example which shows the use of **CellStyleGeneralEnsurer**.

```
public class JxlsTemplatesTest {

    @Test
    public void findMissingCellStyleGeneral() throws IOException {
        findMissingCellStyleGeneral("..");
    }

    private void findMissingCellStyleGeneral(String dir) throws IOException {
        List<String> errors = new ArrayList<>();
        Files.walkFileTree(Paths.get(dir), new SimpleFileVisitor<Path>() {
            // preVisitDirectory ...

            @Override
            public FileVisitResult visitFile(Path path, BasicFileAttributes attrs) throws IOException {
                if (path.toFile().getName().endsWith(".xlsx")) {
                    check(path, errors);
                }
                return FileVisitResult.CONTINUE;
            }
        });
        if (!errors.isEmpty()) {
            fail(errors.stream().collect(Collectors.joining("\n")));
        }
    }
    
    void check(Path path, List<String> errors) throws IOException {
        try (InputStream in = Files.newInputStream(path)) {
            var g = new CellStyleGeneralEnsurer() {
                @Override
                protected void fix(XSSFCell cell, String content) {
                    errors.add(path.toFile().getName() + " " + cell.getSheet().getSheetName() + "!"
                            + cell.getReference() + "  " + content);
                }
            };
            Workbook workbook = PoiTransformerFactory.openWorkbook(in);
            g.process((XSSFWorkbook) workbook);
            workbook.close();
        }
    }
}
```