import java.awt.Desktop;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class InvoiceMaker {
	private static void printTarget(Path target) {
		Desktop desktop = Desktop.getDesktop();
		try {
			desktop.print(target.toFile());
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static ArrayList<Invoice> getInvoiceList(Path registry, int startPosition, String initiator,
			String paymentDate) throws FileNotFoundException, IOException {
		ArrayList<Invoice> invoiceList = new ArrayList<Invoice>();
		Workbook wb = new XSSFWorkbook(new FileInputStream(registry.toString()));
		Sheet sheet = wb.getSheetAt(0);
		for (int i = 15 + startPosition; i < sheet.getLastRowNum() + 1; i++) {
			Row r = sheet.getRow(i);

			Invoice currentInvoice = new Invoice();

			if (r.getCell(5).getStringCellValue() != "") {
				String department = r.getCell(2).getStringCellValue();
				currentInvoice.setDepartment(department);

				String entity = r.getCell(1).getStringCellValue();
				currentInvoice.setEntity(entity);

				Float invoiceAmount = (float) r.getCell(10).getNumericCellValue();
				currentInvoice.setInvoiceAmount(invoiceAmount);

				// String paymentDate = "22.02.2018";
				currentInvoice.setPaymentDate(paymentDate);

				String contractor = r.getCell(5).getStringCellValue();
				currentInvoice.setContractor(contractor);

				String invoiceNumber = r.getCell(7).getStringCellValue();
				currentInvoice.setInvoiceNumber(invoiceNumber);

				Date parseDate = r.getCell(9).getDateCellValue();
				DateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
				String invoiceDate = dateFormat.format(parseDate);
				currentInvoice.setInvoiceDate(invoiceDate);

				String description = r.getCell(6).getStringCellValue();
				currentInvoice.setDescription(description);

				String branch = r.getCell(4).getStringCellValue();
				currentInvoice.setBranch(branch);

				String item = r.getCell(3).getStringCellValue();
				currentInvoice.setItem(item);

				// String initiator = "Воробьев В.А.";
				currentInvoice.setInitiator(initiator);

				currentInvoice.setRegistryDate();

				String departmentDirector = "Блинов В.М.";
				currentInvoice.setDepartmentDirector(departmentDirector);

				StringBuffer parseRC = new StringBuffer();
				parseRC.append(currentInvoice.getDepartment());
				parseRC.append(".");
				parseRC.append(currentInvoice.getRegistryDate());
				String registryCode = parseRC.toString();
				currentInvoice.setRegistryCode(registryCode);

				Double parseDouble = r.getCell(0).getNumericCellValue();
				Integer registryNumber = parseDouble.intValue();
				currentInvoice.setRegistryNumber(registryNumber.toString());

				String comment = r.getCell(8).getStringCellValue();
				currentInvoice.setComment(comment);

				System.out.println(registryNumber.toString() + " " + entity + " " + department + " " + item + " "
						+ branch + " " + contractor + " " + description + " " + invoiceNumber + " " + invoiceDate + " "
						+ invoiceAmount + " " + comment);
			} else {
				break;
			}
			invoiceList.add(currentInvoice);
		}
		wb.close();
		return invoiceList;
	}
	
	private static void makeInvoices(Path registry, int startPosition, boolean print, String initiator,
			String paymentDate) throws FileNotFoundException, IOException, ParseException {
		ArrayList<Invoice> invoiceList = getInvoiceList(registry, startPosition, initiator, paymentDate);
		Iterator<Invoice> ili = invoiceList.iterator();
		while (ili.hasNext()) {
			Invoice curInvoice = ili.next();
			Path template = FileSystems.getDefault().getPath("d:", "template.xlsx");
			curInvoice.exportInvoiceToXls(template);
			if (print)
				printTarget(curInvoice.getFilePath());
		}
	}
	
	static String getLoggedInFullName() {
		String fullName;
		fullName = System.getProperty("user.fullname");
		
	    System.out.println("username = " + fullName);
	    
		return fullName;
	}

	public static void main(String[] args) {
		getLoggedInFullName();
/*		Options opt = new Options();

		opt.addOption("help", false, "Help message.");
		opt.addOption("i", true, "Target path for registry file.");
		opt.addOption("s", true, "Starting position in file.");
		opt.addOption("p", false, "Print invoices.");
		opt.addOption("n", true, "Initiator name.");
		opt.addOption("d", true, "Planned payment date.");

		HelpFormatter formatter = new HelpFormatter();

		CommandLineParser p = new DefaultParser();
		try {
			if (opt.hasOption("i") && opt.hasOption("s") && opt.hasOption("n") && opt.hasOption("d")) {
				CommandLine c = p.parse(opt, args);
				Path registry = FileSystems.getDefault().getPath(c.getOptionValue("i"));
				int startPosition = Integer.parseInt(c.getOptionValue("s"));
				makeInvoices(registry, startPosition, c.hasOption("p"), c.getOptionValue("n"), c.getOptionValue("d"));
			} else {
				formatter.printHelp("Not enough options specified", opt);
			}
		} catch (org.apache.commons.cli.ParseException | IOException | ParseException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
*/	}
}