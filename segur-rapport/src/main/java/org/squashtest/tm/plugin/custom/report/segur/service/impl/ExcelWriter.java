/*
 * Copyright ANS 2020-2022
 */
package org.squashtest.tm.plugin.custom.report.segur.service.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.squashtest.tm.plugin.custom.report.segur.Constantes;
import org.squashtest.tm.plugin.custom.report.segur.Level;
import org.squashtest.tm.plugin.custom.report.segur.Message;
import org.squashtest.tm.plugin.custom.report.segur.Parser;
import org.squashtest.tm.plugin.custom.report.segur.Traceur;
import org.squashtest.tm.plugin.custom.report.segur.model.ExcelRow;
import org.squashtest.tm.plugin.custom.report.segur.model.ReqStepBinding;
import org.squashtest.tm.plugin.custom.report.segur.model.Step;
import org.squashtest.tm.plugin.custom.report.segur.model.TestCase;

/**
 * The Class ExcelWriter.
 */
@Component
public class ExcelWriter {

	private static final String EXCLUDED_TC_STATUS = "OBSOLETE";

	private static final String REQ_CONTEXT_PATH = "%srequirement-workspace/requirement/%d/content";

	private static final String TESTCASE_CONTEXT_PATH = "%stest-case-workspace/test-case/%d/content";

	private static final int MAX_STEPS = 10;

	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelWriter.class);

	/** The Constant REM_SHEET_INDEX. */
	// onglets
	public static final int REM_SHEET_INDEX = 0;

	/** The Constant METIER_SHEET_INDEX. */
	public static final int METIER_SHEET_INDEX = 1;

	/** The Constant ERROR_SHEET_NAME. */
	// public static final int ERROR_SHEET_INDEX = 2;
	public static final String ERROR_SHEET_NAME = "WARNING-ERROR";

	/** The Constant REM_FIRST_EMPTY_LINE. */
	// onglet 0
	public static final int REM_FIRST_EMPTY_LINE = 2; // 0-based index '2' <=> line 3

	/** The Constant REM_LINE_STYLE_TEMPLATE_INDEX. */
	public static final int REM_LINE_STYLE_TEMPLATE_INDEX = 1;

	/** The Constant REM_COLUMN_CONDITIONNELLE. */
	public static final int REM_COLUMN_CONDITIONNELLE = 0;



	/** The Constant REM_COLUMN_ID_SECTION. */
	public static final int REM_COLUMN_ID_SECTION = 2;

	/** The Constant REM_COLUMN_SECTION. */
	public static final int REM_COLUMN_SECTION = 3;

	/** The Constant REM_COLUMN_BLOC. */
	public static final int REM_COLUMN_BLOC = 4;
	
	/** The Constant REM_COLUMN_PERIMETRE. */
	public static final int REM_COLUMN_PERIMETRE = 0;
	/** The Constant REM_COLUMN_NUMERO_EXIGENCE. */
	//public static final int REM_COLUMN_NUMERO_EXIGENCE = 7;
	public static final int REM_COLUMN_NUMERO_EXIGENCE = 1; //0;
	/** The Constant REM_COLUMN_CHAPITRE. */
	public static final int REM_COLUMN_CHAPITRE = 2; //1;
	/** The Constant REM_COLUMN_FONCTION. */
	//public static final int REM_COLUMN_FONCTION = 5;
	public static final int REM_COLUMN_FONCTION = 3; //2;
	/** The Constant REM_COLUMN_ENONCE. */
	//public static final int REM_COLUMN_ENONCE = 8;
	public static final int REM_COLUMN_ENONCE = 4; // 3;
	/** The Constant REM_COLUMN_NATURE. */
	//public static final int REM_COLUMN_NATURE = 6;
	public static final int REM_COLUMN_NATURE = 6; //4;
	/** The Constant REM_COLUMN_PROFIL. */
	//public static final int REM_COLUMN_PROFIL = 1;
	public static final int REM_COLUMN_PROFIL = 7; //5;
	/** The Constant REM_COLUMN_NUMERO_SCENARIO. */
	//public static final int REM_COLUMN_NUMERO_SCENARIO = 9;
	public static final int REM_COLUMN_NUMERO_SCENARIO = 8; //6;
	/** The Constant REM_COLUMN_SCENARIO_CONFORMITE. */
	//public static final int REM_COLUMN_SCENARIO_CONFORMITE = 10;
	public static final int REM_COLUMN_SCENARIO_CONFORMITE = 9; // 7;
	
	/** The Constant MAX_STEP_NUMBER. */
	public static final int MAX_STEP_NUMBER = 10;

	/** The Constant REM_COLUMN_FIRST_NUMERO_PREUVE. */
	public static final int REM_COLUMN_FIRST_NUMERO_PREUVE = REM_COLUMN_SCENARIO_CONFORMITE + 1;

    public static final int REM_COLUMN_COMMENTAIRE = 5; //REM_COLUMN_SCENARIO_CONFORMITE + MAX_STEP_NUMBER * 2 + 1;

	public static final int REM_COLUMN_STATUT_PUBLICATION = REM_COLUMN_SCENARIO_CONFORMITE + MAX_STEP_NUMBER * 2 + 1;//REM_COLUMN_COMMENTAIRE + 1;
	
	public static final int REM_COLUMN_VA_NMOINS1 = REM_COLUMN_STATUT_PUBLICATION + 1;

	public static final int REM_COLUMN_CRITICITY = REM_COLUMN_VA_NMOINS1 + 1;

	public static final int REM_COLUMN_ORDRE = REM_COLUMN_CRITICITY + 1;
	/** The Constant PREPUB_COLUMN_BON_POUR_PUBLICATION. */
	//public static final int PREPUB_COLUMN_BON_POUR_PUBLICATION = REM_COLUMN_SCENARIO_CONFORMITE + MAX_STEP_NUMBER * 2
	//		+ 1;
	public static final int PREPUB_COLUMN_BON_POUR_PUBLICATION = REM_COLUMN_SCENARIO_CONFORMITE + MAX_STEP_NUMBER * 2
			+ 5;

	/** The Constant PREPUB_COLUMN_REFERENCE_EXIGENCE. */
	public static final int PREPUB_COLUMN_REFERENCE_EXIGENCE = PREPUB_COLUMN_BON_POUR_PUBLICATION + 1;

	/** The Constant PREPUB_COLUMN_REFERENCE_CAS_DE_TEST. */
	public static final int PREPUB_COLUMN_REFERENCE_CAS_DE_TEST = PREPUB_COLUMN_REFERENCE_EXIGENCE + 1;

	/** The Constant PREPUB_COLUMN_REFERENCE_EXIGENCE_SOCLE. */
	public static final int PREPUB_COLUMN_REFERENCE_EXIGENCE_SOCLE = PREPUB_COLUMN_REFERENCE_CAS_DE_TEST + 1;

	/** The Constant PREPUB_COLUMN_POINTS_DE_VERIF. */
	public static final int PREPUB_COLUMN_POINTS_DE_VERIF = PREPUB_COLUMN_REFERENCE_EXIGENCE_SOCLE + 1;

	/** The Constant PREPUB_COLUMN_NOTE_INTERNE. */
	public static final int PREPUB_COLUMN_NOTE_INTERNE = PREPUB_COLUMN_POINTS_DE_VERIF + 1;

	/** The Constant PREPUB_COLUMN_SEGUR_REM. */
	public static final int PREPUB_COLUMN_SEGUR_REM = PREPUB_COLUMN_NOTE_INTERNE + 1;

	/** The Constant ERROR_COLUMN_LEVEL. */

	public static final int ERROR_COLUMN_LEVEL = 0;

	/** The Constant ERROR_COLUMN_RESID. */
	public static final int ERROR_COLUMN_RESID = 1;

	/** The Constant ERROR_COLUMN_MSG. */
	public static final int ERROR_COLUMN_MSG = 2;

	private Traceur traceur;

	private String squashBaseUrl;

	/**
	 * Instantiates a new excel writer.
	 *
	 * @param traceur the traceur
	 */
	public ExcelWriter(Traceur traceur) {
		super();
		this.traceur = traceur;
	}

	/**
	 * Load workbook template.
	 *
	 * @param templateName the template name
	 * @return the XSSF workbook
	 */
	public XSSFWorkbook loadWorkbookTemplate(String templateName) {

		InputStream template = null;
		XSSFWorkbook wk = null;
		try {
			template = Thread.currentThread().getContextClassLoader().getResourceAsStream(templateName);
			wk = new XSSFWorkbook(template);
			template.close();
		} catch (IOException e) {
			LOGGER.error(" erreur sur création du workbook ... ", e);
		}
		return wk;
	}

	/**
	 * Put datas in workbook.
	 *
	 * @param boolPrebub the bool prebub
	 * @param workbook   the workbook
	 * @param data       the data
	 */
	public void putDatasInWorkbook(XSSFWorkbook workbook, DSRData data) {
		squashBaseUrl = data.getPerimeter().getSquashBaseUrl();
		// Get first sheet
		XSSFSheet sheet = workbook.getSheet("Exigences");
		// Récupération de la ligne 2 pour utilisation des styles
		Row style2apply = sheet.getRow(REM_LINE_STYLE_TEMPLATE_INDEX);
		// ecriture des données
		int lineNumber = REM_FIRST_EMPTY_LINE;
		// Style links
		CreationHelper helper = workbook.getCreationHelper();
		short height = 200;
		Font linkFont = workbook.createFont();
		linkFont.setFontHeight(height);
		linkFont.setFontName("ARIAL");
		linkFont.setUnderline(XSSFFont.U_SINGLE);
		linkFont.setColor(HSSFColorPredefined.BLUE.getIndex());
		// boucle sur les exigences
		for (ExcelRow req : data.getRequirements()) {

			//##DEBUG
			//String reqResume = "Num exi : " + req.getNumeroExigence_8() + " // nat exigence = " + req.getNatureExigence_7()+ " // reference = " + req.getReference()+ " // referenceSocle = " + req.getReferenceSocle();
			//traceur.addMessage(Level.INFO, req.getResId(),reqResume);
			//
			
	
			// extraire les CTs liés à l'exigence de la map du binding
			List<ReqStepBinding> bindingCT = data.getBindings()
					.stream()
					.filter(p -> p.getResId().equals(req.getResId()))
					.distinct().collect(Collectors.toList());
			// Traitement des cas de test 
			
			//Liste finale des cas de test à exporter
			List<Long> tcIds;
			
			// Il faut déterminer quels sont les cas de test à conserver :
			int tcNumberFromREM = bindingCT.stream().filter(b -> b.getFromSocle().equals(Boolean.FALSE))
					.map(item -> item.getTclnId()).collect(Collectors.toList()).size();
			// - Si présence de cas de test dérivés => On conserve uniquement ceux dérivés
			if(tcNumberFromREM > 0) {
				tcIds = bindingCT.stream().filter(b -> b.getFromSocle().equals(Boolean.FALSE))
						.map(item -> item.getTclnId()).collect(Collectors.toList());
			// - Sinon => On conserve tout
			}else {
				tcIds = bindingCT.stream().map(item -> item.getTclnId()).collect(Collectors.toList());
			}

			
			if (tcIds.isEmpty()) {
				// On ajoute un test vide pour pouvoir formater les cellules
				tcIds.add(0L);
			}
			
			//construction d'une liste de testcase à partie de la liste des ID des testcases
			List<TestCase> tcList = new ArrayList<TestCase>();
			for (Long tcID : tcIds) {
				TestCase testCase;
				if (tcID == 0L) {
					testCase = createDummyTestCase(tcID);
				} else {
					if (EXCLUDED_TC_STATUS.equals(data.getTestCases().get(tcID).getTcStatus())) {
						testCase = createDummyTestCase(0L);
					} else {
						testCase = data.getTestCases().get(tcID);
					}
				}
				tcList.add(testCase);
			}

			Collections.sort(tcList);

			for (TestCase testCase : tcList ) {

			// for (Long tcID : tcIds) {
			// 	TestCase testCase;
			// 	if (tcID == 0L) {
			// 		testCase = createDummyTestCase(0L);
			// 	} else {
			// 		if (EXCLUDED_TC_STATUS.equals(data.getTestCases().get(tcID).getTcStatus())) {
			// 			testCase = createDummyTestCase(0L);
			// 		} else {
			// 			testCase = data.getTestCases().get(tcID);
			// 		}
			// 	}
				// on ecrit (ou réecrit) les colonnes sur les exigences
				Row rowWithTC = writeRequirementRow(req, sheet, lineNumber, style2apply);

				// liste des steps pour l'exigence ET le cas de test courant
				if (testCase.getIsCoeurDeMetier()) {

					// TODO onglet coeur de métier => lecture des steps dans le binding
//				bindingSteps = liste.stream().filter(p -> p.getResId().equals(req.getResId()))
//						.map(val -> val.getStepId()).distinct().collect(Collectors.toList());
					writeCaseTestPartCoeurDeMetier(testCase, null, data.getSteps(), rowWithTC, style2apply);
				} else { // non coeur de métier => on prends tous les steps du CT
					writeCaseTestPart(testCase, data.getSteps(), rowWithTC, style2apply);
				}

				// colonnes prépublication nécessitant le CT
				if (data.getPerimeter().isPrePublication()) {
					Cell c32 = rowWithTC.createCell(PREPUB_COLUMN_BON_POUR_PUBLICATION);
					c32.setCellStyle(style2apply.getCell(PREPUB_COLUMN_BON_POUR_PUBLICATION).getCellStyle());
					if (req.getReqStatus().equals(Constantes.STATUS_APPROVED)) {
						ExcelRow socleReq = data.getRequirementById(req.getSocleReqId());
						String socleStatus = "";
						if (socleReq != null) {
							socleStatus = socleReq.getReqStatus();
						}
						if (testCase.getTcln_id() == 0L || socleStatus.equals(Constantes.STATUS_APPROVED)) {
							c32.setCellValue(" X ");
						} else {
							c32.setCellValue(" ");
						}
					}

					Cell c33 = rowWithTC.createCell(PREPUB_COLUMN_REFERENCE_EXIGENCE);
					CellStyle c33Style = rowWithTC.getSheet().getWorkbook().createCellStyle();
					c33Style.cloneStyleFrom(style2apply.getCell(PREPUB_COLUMN_REFERENCE_EXIGENCE).getCellStyle());
					c33Style.setFont(linkFont);
					c33.setCellStyle(c33Style);
					c33.setCellValue(req.getReference());
					XSSFHyperlink c33link = (XSSFHyperlink) helper.createHyperlink(HyperlinkType.URL);
					c33link.setAddress(String.format(REQ_CONTEXT_PATH, squashBaseUrl, req.getReqId()));
					c33.setHyperlink(c33link);

					CellStyle c34Style = rowWithTC.getSheet().getWorkbook().createCellStyle();
					c34Style.cloneStyleFrom(style2apply.getCell(PREPUB_COLUMN_REFERENCE_CAS_DE_TEST).getCellStyle());
					Cell c34 = rowWithTC.createCell(PREPUB_COLUMN_REFERENCE_CAS_DE_TEST);
					c34Style.setFont(linkFont);
					c34.setCellStyle(c34Style);
					if (testCase.getTcln_id() > 0) {
						c34.setCellValue(testCase.getReference());
						XSSFHyperlink c34link = (XSSFHyperlink) helper.createHyperlink(HyperlinkType.URL);
						c34link.setAddress(String.format(TESTCASE_CONTEXT_PATH, squashBaseUrl, testCase.getTcln_id()));
						c34.setHyperlink(c34link);
					}

					Cell c35 = rowWithTC.createCell(PREPUB_COLUMN_REFERENCE_EXIGENCE_SOCLE);
					CellStyle c35Style = rowWithTC.getSheet().getWorkbook().createCellStyle();
					c35Style.cloneStyleFrom(style2apply.getCell(PREPUB_COLUMN_REFERENCE_EXIGENCE_SOCLE).getCellStyle());
					c35Style.setFont(linkFont);
					c35.setCellStyle(c35Style);
					if (req.getSocleReqId() > 0) {
						c35.setCellValue(req.getReferenceSocle());
						XSSFHyperlink c35link = (XSSFHyperlink) helper.createHyperlink(HyperlinkType.URL);
						c35link.setAddress(String.format(REQ_CONTEXT_PATH, squashBaseUrl, req.getSocleReqId()));
						c35.setHyperlink(c35link);
					}

					Cell c36 = rowWithTC.createCell(PREPUB_COLUMN_POINTS_DE_VERIF);
					CellStyle c36Style = rowWithTC.getSheet().getWorkbook().createCellStyle();
					c36Style.cloneStyleFrom(style2apply.getCell(PREPUB_COLUMN_POINTS_DE_VERIF).getCellStyle());
					c36.setCellStyle(c36Style);
					c36.setCellValue(testCase.getPointsDeVerification());

					Cell c37 = rowWithTC.createCell(PREPUB_COLUMN_NOTE_INTERNE);
					CellStyle c37Style = rowWithTC.getSheet().getWorkbook().createCellStyle();
					c37Style.cloneStyleFrom(style2apply.getCell(PREPUB_COLUMN_NOTE_INTERNE).getCellStyle());
					c37.setCellStyle(c37Style);
					c37.setCellValue(Parser.convertHTMLtoString(req.getNoteInterne()));

					Cell c38 = rowWithTC.createCell(PREPUB_COLUMN_SEGUR_REM);
					CellStyle c38Style = rowWithTC.getSheet().getWorkbook().createCellStyle();
					c38Style.cloneStyleFrom(style2apply.getCell(PREPUB_COLUMN_SEGUR_REM).getCellStyle());
					c38.setCellStyle(c38Style);
					c38.setCellValue(req.getSegurRem());

				}
				lineNumber++;
			}

		} // exigences
			// Suppression de la ligne 1 (template de style)
		sheet.shiftRows(REM_LINE_STYLE_TEMPLATE_INDEX + 1, lineNumber - 1, -1);
		// add borders to cells
		for (Row row : sheet) {
			for (Cell cell : row) {
				CellStyle style = cell.getCellStyle();
				style.setBorderBottom(BorderStyle.THIN);
				style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				style.setBorderLeft(BorderStyle.THIN);
				style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				style.setBorderRight(BorderStyle.THIN);
				style.setRightBorderColor(IndexedColors.BLACK.getIndex());
				style.setBorderTop(BorderStyle.THIN);
				style.setTopBorderColor(IndexedColors.BLACK.getIndex());
				cell.setCellStyle(style);
			}
		}

		writeErrorSheet(workbook);

		LOGGER.info("  fin remplissage du woorkbook: " + workbook);

		if (!data.getPerimeter().isPrePublication()) {
			lockWorkbook(workbook);
		}
	}

	/**
	 * Flush to temporary file.
	 *
	 * @param workbook the workbook
	 * @param filename the file name
	 * @return the file
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public File flushToTemporaryFile(XSSFWorkbook workbook, String filename) throws IOException {
		String tmpdir = System.getProperty("java.io.tmpdir");
		String absolutePath = tmpdir + File.separator + filename;
		File tempFile = new File(absolutePath);
		tempFile.delete();
		FileOutputStream out = new FileOutputStream(tempFile);
		workbook.write(out);
		workbook.close();
		out.close();
		return tempFile;
	}

	private Row writeRequirementRow(ExcelRow data, XSSFSheet sheet, int lineIndex, Row style2apply) {
		// ecriture des données

		Row row = sheet.createRow(lineIndex);

		Cell c0 = row.createCell(REM_COLUMN_PERIMETRE);
		CellStyle c0Style = sheet.getWorkbook().createCellStyle();
		c0Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_PERIMETRE).getCellStyle());
		c0.setCellStyle(c0Style);
		c0.setCellValue(data.getPerimetre_10());

		Cell c1 = row.createCell(REM_COLUMN_PROFIL);
		CellStyle c1Style = sheet.getWorkbook().createCellStyle();
		c1Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_PROFIL).getCellStyle());
		c1.setCellStyle(c1Style);
		c1.setCellValue(data.getProfil_2());

		Cell c3 = row.createCell(REM_COLUMN_CHAPITRE);
		CellStyle c3Style = sheet.getWorkbook().createCellStyle();
		c3Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_CHAPITRE).getCellStyle());
		c3.setCellStyle(c3Style);
		//c3.setCellValue(data.getSection_4() + " " + data.getBloc_5());
		c3.setCellValue(data.getBloc_5());


 		Cell c6 = row.createCell(REM_COLUMN_FONCTION);
		CellStyle c6Style = sheet.getWorkbook().createCellStyle();
		c6Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_FONCTION).getCellStyle());
		c6.setCellStyle(c6Style);
		c6.setCellValue(data.getFonction_6());

		Cell c7 = row.createCell(REM_COLUMN_NATURE);
		CellStyle c7Style = sheet.getWorkbook().createCellStyle();
		c7Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_NATURE).getCellStyle());
		c7.setCellStyle(c7Style);
		c7.setCellValue(data.getNatureExigence_7());

		Cell c8 = row.createCell(REM_COLUMN_NUMERO_EXIGENCE);
		CellStyle c8Style = sheet.getWorkbook().createCellStyle();
		c8Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_NUMERO_EXIGENCE).getCellStyle());
		c8.setCellStyle(c8Style);
		if (data.getReferenceSocle().isEmpty()) {
			c8.setCellValue(extractNumberFromReference(data.getNumeroExigence_8()));
		} else {
			c8.setCellValue(data.getNumeroExigence_8());
		}

		Cell c9 = row.createCell(REM_COLUMN_ENONCE);
		CellStyle c9Style = sheet.getWorkbook().createCellStyle();
		c9Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_ENONCE).getCellStyle());
		c9.setCellStyle(c9Style);
		c9Style.setWrapText(true);
		c9.setCellValue(Parser.convertHTMLtoString(data.getEnonceExigence_9()));

		Cell c12 = row.createCell(REM_COLUMN_COMMENTAIRE);
		CellStyle c12Style = sheet.getWorkbook().createCellStyle();
		c12Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_COMMENTAIRE).getCellStyle());
		c12.setCellStyle(c12Style);
		c12Style.setWrapText(true);
		c12.setCellValue(Parser.convertHTMLtoString(data.getCommentaire())); 
		
		Cell c13 = row.createCell(REM_COLUMN_STATUT_PUBLICATION);
		CellStyle c13Style = sheet.getWorkbook().createCellStyle();
		c13Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_STATUT_PUBLICATION).getCellStyle());
		c13.setCellStyle(c13Style);
		c13Style.setWrapText(true);
		c13.setCellValue(data.getStatutPublication()); 

	 	Cell c14 = row.createCell(REM_COLUMN_VA_NMOINS1);
		CellStyle c14Style = sheet.getWorkbook().createCellStyle();
		c14Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_VA_NMOINS1).getCellStyle());
		c14.setCellStyle(c14Style);
		c14Style.setWrapText(true);
		c14.setCellValue(data.getVaNMoins1());  

		Cell c15 = row.createCell(REM_COLUMN_CRITICITY);
		CellStyle c15Style = sheet.getWorkbook().createCellStyle();
		c15Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_CRITICITY).getCellStyle());
		c15.setCellStyle(c15Style);
		c15Style.setWrapText(true);
		c15.setCellValue(data.getCriticite_11());  

		Cell c16 = row.createCell(REM_COLUMN_ORDRE);
		CellStyle c16Style = sheet.getWorkbook().createCellStyle();
		c16Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_ORDRE).getCellStyle());
		c16.setCellStyle(c16Style);
		c16Style.setWrapText(true);
		c16.setCellValue(data.getOrdre());  

		return row;

	}

	private void writeCaseTestPart(TestCase testcase, Map<Long, Step> steps, Row row, Row style2apply) {
		// ecriture des données
		CellStyle c10Style = row.getSheet().getWorkbook().createCellStyle();
		c10Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_NUMERO_SCENARIO).getCellStyle());
		Cell c10 = row.createCell(REM_COLUMN_NUMERO_SCENARIO);
		c10.setCellStyle(c10Style);
		c10.setCellValue(extractNumberFromReference(testcase.getReference()));

		// cas des CTs non coeur de métier
		Cell c11 = row.createCell(REM_COLUMN_SCENARIO_CONFORMITE);
		CellStyle c11Style = row.getSheet().getWorkbook().createCellStyle();
		c11Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_SCENARIO_CONFORMITE).getCellStyle());
		c11.setCellStyle(c11Style);
		if (testcase.getTcln_id() > 0) {
			String content = "";
			//if (!"".equals(testcase.getPrerequisite())) {
			//	content += "Prérequis : " + Parser.convertHTMLtoString(testcase.getPrerequisite());
			//	if (!"".equals(testcase.getDescription())) {
			//		content += Constantes.LINE_SEPARATOR
			//				+ Parser.convertHTMLtoString(testcase.getDescription());
			//	}
			//} else { // cas où une description existe sans prérequis
			//	if (!"".equals(testcase.getDescription())) {
			//		content += Parser.convertHTMLtoString(testcase.getDescription());
			//	}
			//}
			if (!"".equals(testcase.getDescription())) {
					content += Parser.convertHTMLtoString(testcase.getDescription());
			}
			c11.setCellValue(content);
		}
		// les steps sont reordonnées dans la liste à partir de leur référence
		int currentExcelColumn = REM_COLUMN_FIRST_NUMERO_PREUVE;
		List<Step> testSteps = new ArrayList<>();
		if (testcase.getOrderedStepIds() != null) {
			for (Long id : testcase.getOrderedStepIds()) {
				testSteps.add(steps.get(id));
			}
		}
		Collections.sort(testSteps);
		if (testSteps.size() < MAX_STEPS) {
			for (int i = testSteps.size(); i < MAX_STEPS; i++) {
				testSteps.add(new Step(Long.valueOf(i), "", i));
			}
		}
		for (Step step : testSteps) {
			if (currentExcelColumn > REM_COLUMN_FIRST_NUMERO_PREUVE + MAX_STEPS * 2) {
				traceur.addMessage(Level.WARNING, testcase.getTcln_id(),
						String.format("Le test contient plus de %s preuves", MAX_STEPS));
				break;
			}
			Cell c12plus = row.createCell(currentExcelColumn);
			CellStyle c12Style = row.getSheet().getWorkbook().createCellStyle();
			c12Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_FIRST_NUMERO_PREUVE).getCellStyle());
			c12plus.setCellStyle(c12Style);
			c12plus.setCellValue(extractNumberFromReference(step.getReference()));
			currentExcelColumn++;

			Cell resultCell = row.createCell(currentExcelColumn);
			CellStyle c13Style = row.getSheet().getWorkbook().createCellStyle();
			c13Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_FIRST_NUMERO_PREUVE + 1).getCellStyle());
			c13Style.setWrapText(true);
			resultCell.setCellStyle(c13Style);
			resultCell.setCellValue(Parser.convertHTMLtoString(step.getExpectedResult()));
			currentExcelColumn++;
		}
	}

	private void writeCaseTestPartCoeurDeMetier(TestCase testcase, List<Long> bindedStepIds, Map<Long, Step> steps,
			Row row, Row style2apply) {
		CellStyle c10Style = row.getSheet().getWorkbook().createCellStyle();
		c10Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_NUMERO_SCENARIO).getCellStyle());
		Cell c10 = row.createCell(REM_COLUMN_NUMERO_SCENARIO);
		c10.setCellStyle(c10Style);
		c10.setCellValue(testcase.getReference());

		// cas des CTs coeur de métier
		Cell c11 = row.createCell(REM_COLUMN_SCENARIO_CONFORMITE);
		CellStyle c11Style = row.getSheet().getWorkbook().createCellStyle();
		c11Style.cloneStyleFrom(style2apply.getCell(REM_COLUMN_SCENARIO_CONFORMITE).getCellStyle());
		c11.setCellStyle(c11Style);
		c11.setCellValue(String.format("Cf. Scénarios Coeur de métier\n%s\n[%d] preuve(s)", testcase.getDescription(),
				testcase.getOrderedStepIds().size()));
		// Création de cellules vides pour chaque step afin de respecter le formatage
		for (int i = REM_COLUMN_SCENARIO_CONFORMITE + 1; i <= REM_COLUMN_SCENARIO_CONFORMITE + MAX_STEPS * 2; i++) {
			Cell blank = row.createCell(i);
			blank.setCellStyle(row.getSheet().getWorkbook().createCellStyle());
		}
	}

	private void writeErrorSheet(XSSFWorkbook workbook) {
		List<Message> msg = traceur.getMsg();
		if (msg.size() != 0) {
			XSSFSheet errorSheet = workbook.createSheet(ERROR_SHEET_NAME);
			int line = 0;
			Row firstRow = errorSheet.createRow(line);
			firstRow.createCell(ERROR_COLUMN_MSG).setCellValue(
					"ATTENTION, le nombre maximum d'erreurs/warnings affichés est : " + Traceur.getMAX_MSG());

			for (Message msgLine : msg) {
				Row row = errorSheet.createRow(++line);
				LOGGER.info("  C1: " + msgLine.getLevel().name());
				row.createCell(ERROR_COLUMN_LEVEL).setCellValue(msgLine.getLevel().name());
				LOGGER.info("  C2 : " + msgLine.getResId());
				row.createCell(ERROR_COLUMN_RESID).setCellValue(msgLine.getResId());
				LOGGER.info("  C3: " + msgLine.getMsg());
				row.createCell(ERROR_COLUMN_MSG).setCellValue(msgLine.getMsg());
			}
		}
	}

	private String extractNumberFromReference(String reference) {
		// supprime le prefix SC, CH, XXX
		String numero = "";
		String prefix = "";
		if (reference != null) {
			int separator = reference.indexOf(".");
			if (separator >= 1) {
				prefix = reference.substring(0, separator);
			}

			if ((prefix.equals(Constantes.PREFIX_PROJET_SOCLE)) || (prefix.equals(Constantes.PREFIX_PROJET_CHANTIER))
					|| (prefix.length() == Constantes.PREFIX_PROJET__METIER_SIZE)) {
				numero = reference.substring(separator + 1, reference.length());
			} else {
				traceur.addMessage(Level.ERROR, reference,
						"Calcul du numéro à partir de la référence : erreur sur suppression du prefix de l'item (ni SC., ni CH., ni XXX.)");
				numero = reference;
			}
		}
		return numero;
	}

	private void lockWorkbook(XSSFWorkbook workbook) {
		LOGGER.info("Appel pour lock d'une feuille du workbook");

		XSSFSheet requirementsSheet = workbook.getSheet("Exigences");
		if (requirementsSheet != null) {
			lockSheet(requirementsSheet);
		}

		XSSFSheet testCasesSheet = workbook.getSheet("Scénarios coeur de métier");
		if (testCasesSheet != null) {
			lockSheet(testCasesSheet);
		}

		workbook.lockStructure();

		workbook.lockRevision();

	}

	private void lockSheet(XSSFSheet sheet) {
		sheet.lockDeleteRows(true);
		sheet.lockDeleteColumns(true);
		sheet.lockInsertColumns(true);
		sheet.lockInsertRows(true);
		sheet.lockSort(false);
		sheet.lockFormatCells(false);
		sheet.lockFormatColumns(false);
		sheet.lockFormatRows(false);
		sheet.lockAutoFilter(false);
		sheet.lockInsertHyperlinks(false);
		String password = generateRandomPassword();
		LOGGER.info("Unlock Password : {}", password);
		sheet.protectSheet(password);
		sheet.enableLocking();
	}

	private TestCase createDummyTestCase(Long tcID) {
		TestCase testCase = new TestCase(tcID, "", "", "", "", "");
		List<Long> stepIds = new ArrayList<Long>();
		testCase.setOrderedStepIds(stepIds);
		return testCase;
	}

	private String generateRandomPassword() {
		return RandomStringUtils.random(255, 33, 122, false, false);
	}

}
