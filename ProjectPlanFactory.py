from ProjectPlan import ProjectPlan


class ProjectPlanFactory:

    def __init__(self, smart, report_id, mapping_sheet, mapping_sheet_column_map, project_type_deletion_dict,
                 log_id, log_column_map, logger):
        self.smart = smart
        self.Report = self.smart.Reports.get_report(report_id)
        self.mapping_sheet = mapping_sheet
        self.mapping_sheet_column_map = mapping_sheet_column_map
        self.project_type_deletion_dict = project_type_deletion_dict
        self.log_id = log_id
        self.log_column_map = log_column_map
        self.logger = logger

    def generate_project_plans(self):
        project_plan_sheets = []
        for row in self.Report.rows:
            project_plan_sheets.append(row.sheet_id)

        if project_plan_sheets:
            self.logger.info("Deleting rows from {} projects".format(len(project_plan_sheets)))
            for sheet_id in project_plan_sheets:
                project_plan = ProjectPlan(self.smart, sheet_id, self.mapping_sheet,
                                           self.mapping_sheet_column_map, self.project_type_deletion_dict,
                                           self.log_id, self.log_column_map, self.logger)

                project_plan.process_rows()
        else:
            self.logger.info("no projects to delete rows from")


