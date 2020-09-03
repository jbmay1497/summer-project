

class ProjectPlan:

    def __init__(self, smart, sheet_id, mapping_sheet, mapping_sheet_column_map, project_type_deletion_dict,
                 log_id, log_column_map, logger):
        self.smart = smart
        self.sheet_id = sheet_id
        self.sheet = self.smart.Sheets.get_sheet(self.sheet_id)
        self.mapping_sheet = mapping_sheet
        self.sheet_column_map = self._build_column_map()
        self.mapping_sheet_column_map = mapping_sheet_column_map
        self.project_type_deletion_dict = project_type_deletion_dict
        self.log_id = log_id
        self.log_column_map = log_column_map
        self.logger = logger

    def _get_cell_by_column_name(self, row, column_name, column_map):
        column_id = column_map[column_name]
        return row.get_column(column_id)

    def _build_column_map(self):
        sheet_column_map = {}
        for column in self.sheet.columns:
            sheet_column_map[column.title] = column.id
        return sheet_column_map

    def _evaluate_row_and_build_deletes(self, source_row, project_plan_type):
        project_plan_type_to_check = self._get_cell_by_column_name(source_row, "Project Plan Type",
                                                                   self.sheet_column_map).display_value

        if self.project_type_deletion_dict[project_plan_type_to_check] != "N/A" and \
                project_plan_type_to_check != project_plan_type and source_row.parent_id:
            return source_row.parent_id
        return None

    def _delete_rows(self, rows_to_delete):
        result = None
        if rows_to_delete:
            self.logger.info("deleting rows now")

            result = self.smart.Sheets.delete_rows(self.sheet_id, rows_to_delete, ignore_rows_not_found=True)
        return result

    def _update_log(self, request_id, project_plan_type, status):

        if not request_id:
            request_id = ""
        if not project_plan_type:
            project_plan_type = ""
        if not status:
            status = ""

        row = self.smart.models.Row()
        row.to_Top = True

        request_cell = self.smart.models.Cell()
        request_cell.column_id = self.log_column_map["Request ID"]
        request_cell.value = request_id

        row.cells.append(request_cell)

        project_plan_type_cell = self.smart.models.Cell()
        project_plan_type_cell.column_id = self.log_column_map["Project Plan Type"]
        project_plan_type_cell.value = project_plan_type

        row.cells.append(project_plan_type_cell)

        status_cell = self.smart.models.Cell()
        status_cell.column_id = self.log_column_map["Status"]

        if status:
            status_cell.value = "Complete"
        else:
            status_cell.value = "Incomplete"

        row.cells.append(status_cell)

        response = self.smart.Sheets.add_rows(
            self.log_id,
            [row]
        )

        return response

    def _update_checkbox(self, request_id_row):

        updated_checkbox_cell = self.smart.models.Cell()
        updated_checkbox_cell.type = "CHECKBOX"
        updated_checkbox_cell.value = True
        updated_checkbox_cell.column_id = self.sheet_column_map["Code Log Run"]

        new_row = self.smart.models.Row()
        new_row.id = request_id_row.id
        new_row.cells.append(updated_checkbox_cell)

        result = self.smart.Sheets.update_rows(self.sheet_id, [new_row])

    def process_rows(self):
        self.logger.info("Deleting rows from sheet " + self.sheet.name)

        project_type = None
        for row in self.sheet.rows:
            if self._get_cell_by_column_name(row, "Status", self.sheet_column_map).display_value == "Project Type":
                project_type = self._get_cell_by_column_name(row, "Notes", self.sheet_column_map).display_value
                break

        subtype = None
        for row in self.sheet.rows:
            if self._get_cell_by_column_name(row, "Status", self.sheet_column_map).display_value == "Project Sub-Type":
                subtype = self._get_cell_by_column_name(row, "Notes", self.sheet_column_map).display_value
                break

        request_id = None
        request_id_row = None
        for row in self.sheet.rows:

            if self._get_cell_by_column_name(row, "Status", self.sheet_column_map).display_value == "Request ID":
                request_id = self._get_cell_by_column_name(row, "Notes", self.sheet_column_map).display_value
                request_id_row = row
                break

        project_plan_type = None
        status = None
        if project_type and subtype:

            for row in self.mapping_sheet.rows:
                map_project_type = self._get_cell_by_column_name(row, "Project Type",
                                                                 self.mapping_sheet_column_map).display_value
                map_subtype = self._get_cell_by_column_name(row, "Project Sub-Type",
                                                                 self.mapping_sheet_column_map).display_value
                if map_project_type == project_type and map_subtype == subtype:
                    project_plan_type = self._get_cell_by_column_name(row, "Project Plan Type",
                                                                      self.mapping_sheet_column_map).display_value
                    break

            rows_to_delete = []
            for row in self.sheet.rows:
                row_to_delete = self._evaluate_row_and_build_deletes(row, project_plan_type)

                if row_to_delete is not None and not (row_to_delete in rows_to_delete):
                    rows_to_delete.append(row_to_delete)

            results = self._delete_rows(rows_to_delete)
            if results.message == "SUCCESS":
                self.logger.info("Rows deleted successfully")
                status = True
                self.logger.info("checking box")
                self._update_checkbox(request_id_row)
                self.logger.info("finished checking box")
            else:
                self.logger.info("Rows not deleted successfully")
                status = False
        else:
            self.logger.info("Blank/Incorrect project type or project subtype specified")
            status = False

        self.logger.info("beginning log")
        self._update_log(request_id, project_plan_type, status)
        self.logger.info("finishing log")
