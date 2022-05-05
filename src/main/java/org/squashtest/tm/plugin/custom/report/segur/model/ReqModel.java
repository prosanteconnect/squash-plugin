package org.squashtest.tm.plugin.custom.report.segur.model;

import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.squashtest.tm.plugin.custom.report.segur.Constantes;
import org.squashtest.tm.plugin.custom.report.segur.Level;
import org.squashtest.tm.plugin.custom.report.segur.Parser;
import org.squashtest.tm.plugin.custom.report.segur.Traceur;
import org.squashtest.tm.plugin.custom.report.segur.service.impl.ExcelWriterUtil;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReqModel {

	private Long resId;
	private Long projectId;
	private String categorie;
	private String description;
	private String reference;
	private String requirementStatus;

	private ExcelData excelData = new ExcelData();

	private List<Cuf> cufs;

	// calcul�e
	private String idSection;
	private String section;

	private Traceur traceur;

	public ReqModel(Long resId, String reference, String requirementStatus, String categorie, String description) {
		super();
		this.resId = resId;
		this.categorie = categorie;
		this.description = description;
		this.reference = reference;
		this.requirementStatus = requirementStatus;
	}

	public ExcelData updateData(Traceur traceur) {
//			public ExcelData updateData () {
		this.traceur = traceur;

		// id nécessaire pour lecture des liens exigence-CTs-(steps)
		excelData.setResId(resId);
		// les cufs ont �t� lus en BDD, on met � jour "excelData"

		Cuf rawProfil = findSpecificCuf(Constantes.PROFIL);
		if (rawProfil == null) {
			traceur.addMessage(Level.ERROR, resId, "pas de cuf 'PROFIL' trouvé");
		} else {
			calculExigenceConditionelle(rawProfil.getLabel());
		}

		excelData.setProfil_2(rawProfil.getLabel());

		// traitement de la section
		Cuf rawSection = findSpecificCuf(Constantes.SECTION);
		splitSectionAndSetExcelData(rawSection.getLabel());

		excelData.setBloc_5(findSpecificCuf(Constantes.BLOC).getLabel());

		excelData.setFonction_6(findSpecificCuf(Constantes.FONCTION).getLabel());

		calculCategorieNature(categorie);

		excelData.setNumeroExigence_8(reference);

		excelData.setEnonceExigence_9(Parser.convertHTMLtoString(description));

		return excelData;
	}

	public Cuf findSpecificCuf(String cufCode) {
		List<Cuf> found = cufs.stream().filter(currentCuf -> cufCode.equals(currentCuf.getCode()))
				.collect(Collectors.toList());		
		//concatenation si plusieurs tags ...
		String labels = found.stream()
        .map( n -> n.getLabel())
        .collect( Collectors.joining( " , " ) );
		found.get(0).setLabel(labels);
		return found.get(0); 
	}

	public void splitSectionAndSetExcelData(String cufSection) {
		int separator = cufSection.indexOf(Constantes.SECTION_SEPARATOR);
		if (separator == -1) {
			traceur.addMessage(Level.ERROR,resId,  "Impossible d'extraitre d'idSection et la section du CUF Section: " + cufSection);
		} else {
			excelData.setId_section_3(cufSection.substring(0, separator));
			excelData.setSection_4(cufSection.substring(separator + 1));
		}
	}

	public void calculExigenceConditionelle(String labelProfil) {
	
		if (labelProfil.isEmpty()) {
			traceur.addMessage(Level.ERROR, resId , "calculExigenceConditionelle impossible cuf Profil non renseigné ");
			return; //non renseigné par défaut
		}
		
		if (labelProfil.equalsIgnoreCase(Constantes.PROFIL_GENERAL)) {
			excelData.setBoolExigenceConditionnelle_1(Constantes.NON);
		} else {
			excelData.setBoolExigenceConditionnelle_1(Constantes.OUI);
		}
	}

	public void calculCategorieNature(String categorie) {
		categorie = categorie.replace('é', 'e');
		categorie = categorie.replace('É', 'E');
		
		if (categorie.toUpperCase().contains(Constantes.CATEGORIE_EXIGENCE)) {
			excelData.setNatureExigence_7(Constantes.CATEGORIE_EXIGENCE);
		} else if (categorie.toUpperCase().contains(Constantes.CATEGORIE_PRECONISATION)) {
			excelData.setNatureExigence_7(Constantes.CATEGORIE_PRECONISATION);
		} else {
			traceur.addMessage(Level.WARNING, resId, "Impossible d'identifier la nature pour l'exigence. Cuf'Catégorie'= " + 
			categorie);
		}
	}

}
