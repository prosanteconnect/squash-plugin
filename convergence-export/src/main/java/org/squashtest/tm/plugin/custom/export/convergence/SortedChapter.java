package org.squashtest.tm.plugin.custom.export.convergence;

public enum SortedChapter {
    FIRST("test1"), SECOND("Alimentation du DMP via une PFI"), THIRD("Transmission via MS-Santé");

	String value;

	SortedChapter(String value) {
		this.value = value;
	}

	public String getValue() {
		return value;
	}
}
