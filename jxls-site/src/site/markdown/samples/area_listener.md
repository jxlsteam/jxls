Area Listener Example
=====================

Introduction
------------

In this example we will demonstrate how to use *AreaListener* for additional control over Excel generation with Jxls.

We will use *Department* and *Employee* objects

    public class Department {
        private String name;
        private Employee chief;
        private List<Employee> staff = new ArrayList<Employee>();
        private String link;

        // getters/setters
        ...
    }

    public class Employee {
        private String name;
        private int age;
        private Double payment;
        private Double bonus;
        private Date birthDate;
        private Employee superior;

        // getters/setters
        ...
    }

The goal is to highlight those bonus cells which have bonus value more than 20%.

Report template
---------------

The [report template](../xls/each_if_demo_template.xls) for this example uses [Comment markup](../reference/excel_markup.html)  to define the transformation areas.
It looks like this

![AreaListener template](../images/each_if_demo_template.png)

Java code
---------

In this example we will use Jxls POI adapter to generate the report.
The code is like the following

        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        try(InputStream is = EachIfCommandDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                System.out.println("Creating area");
                XlsArea xlsArea = new XlsArea("Template!A1:G15", transformer);
                XlsArea departmentArea = new XlsArea("Template!A2:G12", transformer);
                EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);
                XlsArea employeeArea = new XlsArea("Template!A9:F9", transformer);
                XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
                XlsArea elseArea = new XlsArea("Template!A9:F9", transformer);
                IfCommand ifCommand = new IfCommand("employee.payment <= 2000",
                        ifArea,
                        elseArea);
                ifArea.addAreaListener(new SimpleAreaListener(ifArea));
                elseArea.addAreaListener(new SimpleAreaListener(elseArea));
                employeeArea.addCommand(new AreaRef("Template!A9:F9"), ifCommand);
                Command employeeEachCommand = new EachCommand("employee", "department.staff", employeeArea);
                departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);
                xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);
                Context context = new Context();
                context.putVar("departments", departments);
                logger.info("Applying at cell " + new CellRef("Down!A1"));
                xlsArea.applyAt(new CellRef("Down!A1"), context);
                xlsArea.processFormulas();
                logger.info("Setting EachCommand direction to Right");
                departmentEachCommand.setDirection(EachCommand.Direction.RIGHT);
                logger.info("Applying at cell " + new CellRef("Right!A1"));
                xlsArea.reset();
                xlsArea.applyAt(new CellRef("Right!A1"), context);
                xlsArea.processFormulas();
                logger.info("Complete");
                transformer.write();
                logger.info("written to file");
            }
        }

To highlight required bonus cells we adding our custom *SimpleAreaListener* to each area of *If-Command*

        ifArea.addAreaListener(new SimpleAreaListener(ifArea));
        elseArea.addAreaListener(new SimpleAreaListener(elseArea));

The *SimpleAreaListener* is a simple class that highlights the greater 20% bonus cells.

It looks like this

    public class SimpleAreaListener implements AreaListener {
        static Logger logger = LoggerFactory.getLogger(SimpleAreaListener.class);

        XlsArea area;
        PoiTransformer transformer;
        private final CellRef bonusCell1 = new CellRef("Template!E9");
        private final CellRef bonusCell2 =new CellRef("Template!E18");

        public SimpleAreaListener(XlsArea area) {
            this.area = area;
            transformer = (PoiTransformer) area.getTransformer();
        }

        public void beforeApplyAtCell(CellRef cellRef, Context context) {

        }

        public void afterApplyAtCell(CellRef cellRef, Context context) {

        }

        public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {

        }

        public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
            if(bonusCell1.equals(srcCell) || bonusCell2.equals(srcCell)){ // we are at employee bonus cell
                Employee employee = (Employee) context.getVar("employee");
                if( employee.getBonus() >= 0.2 ){ // highlight bonus when >= 20%
                    logger.info("highlighting bonus for employee " + employee.getName());
                    highlightBonus(targetCell);
                }
            }
        }

        private void highlightBonus(CellRef cellRef) {
            Workbook workbook = transformer.getWorkbook();
            Sheet sheet = workbook.getSheet(cellRef.getSheetName());
            Cell cell = sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
            CellStyle cellStyle = cell.getCellStyle();
            CellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.setDataFormat( cellStyle.getDataFormat() );
            newCellStyle.setFont( workbook.getFontAt( cellStyle.getFontIndex() ));
            newCellStyle.setFillBackgroundColor( cellStyle.getFillBackgroundColor());
            newCellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            newCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            cell.setCellStyle(newCellStyle);
        }
    }

As seen we are overriding `afterTransformCell(CellRef srcCell, CellRef targetCell, Context context)` method and trigger cell highlighting if bonus condition is satisfied

        public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
            if(bonusCell1.equals(srcCell) || bonusCell2.equals(srcCell)){ // we are at employee bonus cell
                Employee employee = (Employee) context.getVar("employee");
                if( employee.getBonus() >= 0.2 ){ // highlight bonus when >= 20%
                    logger.info("highlighting bonus for employee " + employee.getName());
                    highlightBonus(targetCell);
                }
            }
        }

The actual highlight is done using POI API in `highlightBonus(CellRef cellRef)` method

Excel Output
------------

Final report  for this example can be downloaded  [here](../xls/listener_demo_output.xls)

