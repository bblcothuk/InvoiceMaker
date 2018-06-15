import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Invoice {

	private String registryNumber;
	private String registryCode;
	private String entity;
	private String department;
	private String departmentDirector;
	private String initiator;
	private String item;
	private String branch;
	private String contractor;
	private String description;
	private String invoiceNumber;
	private String invoiceDate;
	private Float invoiceAmount;
	private String registryDate;
	private String paymentDate;
	private String comment;
	private Path filePath;
	// private Boolean isPaid;
	// private Date paidDate;

	void setInvoice(String department, String entity, float invoiceAmount, String paymentDate, String contractor,
			String invoiceNumber, String invoiceDate, String description, String branch, String item, String initiator,
			String departmentDirector, String registryCode, String registryNumber) {
		this.setDepartment(department);
		this.setEntity(entity);
		this.setInvoiceAmount(invoiceAmount);
		this.setPaymentDate(paymentDate);
		this.setContractor(contractor);
		this.setInvoiceNumber(invoiceNumber);
		this.setInvoiceDate(invoiceDate);
		this.setDescription(description);
		this.setBranch(branch);
		this.setItem(item);
		this.setInitiator(initiator);
		this.setRegistryDate();
		this.setDepartmentDirector(departmentDirector);
		this.setRegistryCode(registryCode);
		this.setRegistryNumber(registryNumber);
	}

	void exportInvoiceToDB() {
		// Exporting invoice to database.
	}

	void exportInvoiceToXls(Path template) throws FileNotFoundException, IOException, ParseException {
		// Exporting invoice to .xls using template.
		
		Workbook WB = new XSSFWorkbook(new FileInputStream(template.toString()));
		Sheet invoiceSheet = WB.getSheetAt(0);

		Pattern p = Pattern.compile("(::[A-Z|a-z]*::)");

		Iterator<Row> ri = invoiceSheet.rowIterator();
		while (ri.hasNext()) {
			Row curRow = ri.next();
			Iterator<Cell> ic = curRow.cellIterator();
			while (ic.hasNext()) {
				Cell curCell = ic.next();
				if (curCell.getCellTypeEnum() == CellType.STRING) {
					StringBuffer sb = new StringBuffer(curCell.getStringCellValue());
					Matcher m = p.matcher(sb);
					
					while (m.find()) {
						switch (m.group(1)) {
						case "::department::":
							sb = new StringBuffer(m.replaceFirst(this.department));
							curCell.setCellValue(sb.toString());
							break;
						case "::entity::":
							sb = new StringBuffer(m.replaceFirst(this.entity));
							curCell.setCellValue(sb.toString());
							break;
						case "::invoiceAmount::":
							sb = new StringBuffer(m.replaceFirst(this.invoiceAmount.toString()));
							curCell.setCellValue(sb.toString());
							break;
						case "::paymentDate::":
							sb = new StringBuffer(m.replaceFirst(this.paymentDate));
							curCell.setCellValue(sb.toString());
							break;
						case "::contractor::":
							sb = new StringBuffer(m.replaceFirst(this.contractor));
							curCell.setCellValue(sb.toString());
							break;
						case "::invoiceNumber::":
							sb = new StringBuffer(m.replaceFirst(this.invoiceNumber));
							curCell.setCellValue(sb.toString());
							m = p.matcher(sb);
							break;
						case "::invoiceDate::":
							sb = new StringBuffer(m.replaceFirst(this.invoiceDate));
							curCell.setCellValue(sb.toString());
							m = p.matcher(sb);
							break;
						case "::description::":
							sb = new StringBuffer(m.replaceFirst(this.description));
							curCell.setCellValue(sb.toString());
							break;
						case "::branch::":
							sb = new StringBuffer(m.replaceFirst(this.branch));
							curCell.setCellValue(sb.toString());
							break;
						case "::item::":
							sb = new StringBuffer(m.replaceFirst(this.item));
							curCell.setCellValue(sb.toString());
							break;
						case "::initiator::":
							sb = new StringBuffer(m.replaceFirst(this.initiator));
							curCell.setCellValue(sb.toString());
							break;
						case "::registryDate::":
							sb = new StringBuffer(m.replaceFirst(this.registryDate));
							curCell.setCellValue(sb.toString());
							break;
						case "::departmentDirector::":
							sb = new StringBuffer(m.replaceFirst(this.departmentDirector));
							curCell.setCellValue(sb.toString());
							break;
						case "::registryCode::":
							sb = new StringBuffer(m.replaceFirst(this.registryCode));
							curCell.setCellValue(sb.toString());
							m = p.matcher(sb);
							break;
						case "::registryNumber::":
							sb = new StringBuffer(m.replaceFirst(this.registryNumber));
							curCell.setCellValue(sb.toString());
							m = p.matcher(sb);
							break;
						case "::comment::":
							sb = new StringBuffer(m.replaceFirst(this.comment));
							curCell.setCellValue(sb.toString());
							m = p.matcher(sb);
							break;
						default:
							System.out.println("Default case.");
							break;
						}
					}
				}
			}
		}
		
		this.makeFilePath();
		if (!this.getFilePath().getParent().toFile().exists()) this.getFilePath().getParent().toFile().mkdirs();
		
		FileOutputStream os = new FileOutputStream(this.filePath.toString());
		WB.write(os);
		os.close();
		WB.close();
	}

	public String getRegistryNumber() {
		return registryNumber;
	}

	public void setRegistryNumber(String registryNumber) {
		this.registryNumber = registryNumber;
	}

	public String getRegistryCode() {
		return registryCode;
	}

	public void setRegistryCode(String registryCode) {
		this.registryCode = registryCode;
	}

	public String getItem() {
		return item;
	}

	public void setItem(String item) {
		this.item = item;
	}

	public String getBranch() {
		return branch;
	}

	public void setBranch(String branch) {
		this.branch = branch;
	}

	public String getContractor() {
		return contractor;
	}

	public void setContractor(String contractor) {
		this.contractor = contractor;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getInvoiceNumber() {
		return invoiceNumber;
	}

	public void setInvoiceNumber(String invoiceNumber) {
		this.invoiceNumber = invoiceNumber;
	}

	public String getInvoiceDate() {
		return invoiceDate;
	}

	public void setInvoiceDate(String invoiceDate) {
		this.invoiceDate = invoiceDate;
	}

	public Float getInvAmount() {
		return invoiceAmount;
	}

	public void setInvoiceAmount(float invoiceAmount) {
		this.invoiceAmount = invoiceAmount;
	}

	public String getDepartment() {
		return department;
	}

	public void setDepartment(String department) {
		this.department = department;
	}

	public String getEntity() {
		return entity;
	}

	public void setEntity(String entity) {
		this.entity = entity;
	}

	public String getRegistryDate() {
		return registryDate;
	}

	public void setRegistryDate() {
		DateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
		this.registryDate = dateFormat.format(new Date());
	}

	public String getPaymentDate() {
		return paymentDate;
	}

	public void setPaymentDate(String paymentDate) {
		this.paymentDate = paymentDate;
	}

	/*
	 * public Boolean getIsPaid() { return isPaid; }
	 * 
	 * public void setIsPaid(Boolean isPaid, Date paidDate) { if (isPaid ==
	 * true) this.paidDate = paidDate; this.isPaid = isPaid; }
	 * 
	 * public Date getPaidDate() { return paidDate; }
	 */
	public String getDepartmentDirector() {
		return departmentDirector;
	}

	public void setDepartmentDirector(String departmentDirector) {
		this.departmentDirector = departmentDirector;
	}

	public String getInitiator() {
		return initiator;
	}

	public void setInitiator(String initiator) {
		this.initiator = initiator;
	}

	public void generateRegistryDate() {
		/*
		 * DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
		 * 
		 * LocalDate d = LocalDate.now(); d.getDayOfWeek();
		 * 
		 * 
		 * return d.withDayOfWeek(DateTimeConstants.FRIDAY)); } else if
		 * (d.getDayOfWeek() == DateTimeConstants.FRIDAY) { // almost useless
		 * branch, could be merged with the one above return d; } else { return
		 * d.plusWeeks(1).withDayOfWeek(DateTimeConstants.FRIDAY)); }
		 * this.paymentDate = dateFormat.format(d);
		 */

	}

	public String getInvoiceFileName() {
				
		StringBuffer parseIFN = new StringBuffer();
		parseIFN.append(this.getInvoiceNumber());
		parseIFN.append("-");
		parseIFN.append(this.getInvoiceDate());
		parseIFN.append(".xlsx");
		
		
		String invoiceFileName = parseFileSystemSymbols(parseIFN.toString());
		
		return invoiceFileName;
	}

	public String getInvoiceMonth() throws ParseException {
		DateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
		Date invoiceDate = dateFormat.parse(this.invoiceDate);
		DateFormat month =  new SimpleDateFormat("MM MMM");
		String invoiceMonth = month.format(invoiceDate);
		return invoiceMonth;
	}
	
	public String getInvoiceYear() throws ParseException {
		DateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
		Date invoiceDate = dateFormat.parse(this.invoiceDate);
		DateFormat year =  new SimpleDateFormat("yyyy");
		String invoiceYear = year.format(invoiceDate);
		return invoiceYear;
	}

	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

	public Path getFilePath() {
		return filePath;
	}
	
	public void makeFilePath() throws ParseException {
		String separator  = FileSystems.getDefault().getSeparator();
		StringBuffer sbPath = new StringBuffer();
		sbPath.append("d:");
		sbPath.append(separator);
		sbPath.append("Заявки");
		sbPath.append(separator);
		sbPath.append(this.getItem());
		sbPath.append(separator);
		sbPath.append(this.getInvoiceYear());
		sbPath.append(separator);
		sbPath.append(this.getInvoiceMonth());
		sbPath.append(separator);
		sbPath.append(this.getContractor());
		sbPath.append(separator);
		sbPath.append(this.getEntity());
		sbPath.append(separator);
		sbPath.append(parseFileSystemSymbols(this.getInvoiceFileName()));
		
		String sPath = /*parseFileSystemSymbols(*/sbPath.toString()/*)*/;
		System.out.println(sPath);
		
		Path filePath = Paths.get(sPath);
		
		setFilePath(filePath);

	}
	
	private static String parseFileSystemSymbols(String string) {
		Pattern p = Pattern.compile("[ \\/:*?\"<>|]+");
		Matcher m = p.matcher(string);
		string = new String(m.replaceAll("_"));
		return string;
	}
	
	public void setFilePath(Path filePath) {
		this.filePath = filePath;
	}
}