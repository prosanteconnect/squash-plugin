/*
 * Copyright ANS 2020-2022
 */
package org.squashtest.tm.plugin.custom.report.segur;

/**
 * The Enum CriteriaEntry.
 */
public enum CriteriaEntry {

	MILESTONES("milestones"),
	PROJECT_IDS("projects"), 
	PREFIXES("prefixes"),
	SEGUR_SELECTION_MODE("segurSelectionMode"),
	SEGUR_PRINT_OPTIONS("segurPrintOptions");

	String value;

	CriteriaEntry(String value) {
		this.value = value;
	}

	public String getValue() {
		return value;
	}
}
