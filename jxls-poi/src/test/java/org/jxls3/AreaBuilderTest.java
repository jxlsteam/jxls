package org.jxls3;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.area.CommandData;
import org.jxls.area.XlsArea;
import org.jxls.builder.xls.AbstractAreaBuilder;
import org.jxls.builder.xls.AreaCommand;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.entity.Employee;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

public class AreaBuilderTest {

	@Test
	public void test() {
		// Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());

        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        JxlsPoiTemplateFillerBuilder builder = JxlsPoiTemplateFillerBuilder.newInstance()
        		.withExpressionNotation("{{", "}}")
        		.withAreaBuilder(new MarkerAreaBuilder());
		tester.test(data, builder);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Employees");
            assertEquals("Elsa", w.getCellValueAsString(2, 1)); // A2 
            assertEquals("1969-05-30T00:00", w.getCellValueAsLocalDateTime(6, 2).toString()); // B6 
            assertEquals(2500d, w.getCellValueAsDouble(4, 3), 0.005d); // C4
            assertNull(w.getCellValueAsString(7, 1)); // A7
        }
        assertNotNull(builder.getAreaBuilder());
	}
	
	static class MarkerAreaBuilder extends AbstractAreaBuilder {
        private final List<Marker> markers = new ArrayList<>();
		
		@Override
		public List<Area> build(Transformer transformer, boolean clearTemplateCells) {
			transformer.getCommentedCells().forEach(d -> findMarkers(d));
			List<String> markersWithoutEnd = markers.stream().filter(m -> m.getEndCell() == null)
					.map(m -> m.getName() + "> " + m.getStartCell().toString()).toList();
			if (!markersWithoutEnd.isEmpty()) {
				throw new RuntimeException("There are markers without end marker:\n" + markersWithoutEnd.toString());
			}
			return super.build(transformer, clearTemplateCells);
		}

		private void findMarkers(CellData d) {
			String[] lines = d.getCellComment().split("\n");
			for (String line : lines) {
				String t = line.trim();
				if (t.endsWith(">") && t.length() > 1) { // start marker
					String name = t.substring(0, t.length() - 1);
					Optional<Marker> mq = markers.stream().filter(m -> name.equals(m.getName())).findFirst();
					if (mq.isPresent()) {
						throw new RuntimeException(
								"Duplicate marker \"" + name + ">\" in cells " + mq.get().getStartCell() + " and " + d.getCellRef());
					}
					markers.add(new Marker(d.getCellRef(), name));
				} else if (t.startsWith("<") && t.length() > 1) { // end marker
					String name = t.substring(1).trim();
					for (Marker m : markers) {
						if (m.getName().equals(name)) {
							m.setEndCell(d.getCellRef());
							break;
						}
					}
				}
			}
		}
		
		@Override
		protected List<CommandData> buildCommands(Transformer transformer, CellData d, String unused) {
			List<CommandData> ret = new ArrayList<>();
			String[] lines = d.getCellComment().split("\n");
			for (String line : lines) {
				String t = line.trim();
				if ("jx:area".equals(t)) {
					area(d, line, new AreaCommand(), transformer, ret);
				} else if (t.startsWith("jx:each ")) {
					t = t.substring("jx:each ".length()).trim();
					int o = t.indexOf(" in ");
					if (o < 0) {
						throw new RuntimeException("Syntax error in command: " + line + "\ncell: " + d.getCellRef());
					}
					EachCommand command = new EachCommand();
					command.setVar(t.substring(0, o).trim());
					command.setItems(t.substring(o + " in ".length()).trim());
					area(d, line, command, transformer, ret);
				}
			}
			return ret;
		}
		
		private void area(CellData d, String line, Command command, Transformer transformer, List<CommandData> ret) {
			for (Marker m : markers) {
				if (m.getStartCell().equals(d.getCellRef())) {
					CommandData c = new CommandData(new AreaRef(d.getCellRef(), m.getEndCell()), command);
					ret.add(c);
					command.addArea(new XlsArea(c.getAreaRef(), transformer));
					return;
				}
			}
			throw new RuntimeException("No marked area for: " + line);
		}
	}
	
	public static class Marker {
		private final String name;
		private final CellRef startCell;
		private CellRef endCell;

		public Marker(CellRef startCell, String name) {
			this.startCell = startCell;
			this.name = name;
		}

		public String getName() {
			return name;
		}

		public CellRef getStartCell() {
			return startCell;
		}

		public CellRef getEndCell() {
			return endCell;
		}

		public void setEndCell(CellRef endCell) {
			this.endCell = endCell;
		}
	}
}
