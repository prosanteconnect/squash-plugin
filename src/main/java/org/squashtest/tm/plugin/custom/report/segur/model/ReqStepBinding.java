package org.squashtest.tm.plugin.custom.report.segur.model;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReqStepBinding {

	// id du lien avec le CT dans requirement_version_coverage
	Long reqVersionCoverageId;

	// requirement_version_coverage_id.verified_req_version_id
	Long resId;

	// id du CT
	Long tclnId;

	// id du step
	Long stepId;
}
