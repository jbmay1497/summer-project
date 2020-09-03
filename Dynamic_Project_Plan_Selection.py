import smartsheet
import logging
import os.path
from ProjectPlanFactory import ProjectPlanFactory
import errno
import sys

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"
access_token = "srz2b0jnfpfjqj1ok88soe99a0"


def setup_logger(logdir):
    try:
        os.makedirs(logdir)
    except Exception as e:
        if e.errno != errno.EEXIST:
            raise

    abp = os.path.abspath(__file__)
    log_name = os.path.basename(abp)[:-3] + '_' + str(os.getpid()) + '.log'
    log_path = os.path.join(logdir, log_name)
    if os.path.exists(log_path):
        with open(log_path, 'w'):
            pass

    if sys.version_info >= (3, 5):
        logging.basicConfig(filename=log_path,
                            filemode='w',
                            format='%(asctime)s|%(levelname)s|%(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S',
                            level=logging.INFO)
    else:
        logging.basicConfig(filename=log_path,
                            filemode='w',
                            format='%(asctime)s|%(levelname)s|%(message)s',
                            date_format='%Y-%m-%d %H:%M:%S',
                            level=logging.INFO)


def get_cell_by_column_name(row, column_name, column_map):
    column_id = column_map[column_name]
    return row.get_column(column_id)


if __name__ == '__main__':

    logdir = os.path.join(os.getcwd(), 'tmp', 'Logs')
    setup_logger(logdir)
    logger = logging.getLogger()

    logger.info("Starting ...")

    # Initialize client
    smart = smartsheet.Smartsheet(access_token)

    # Make sure we don't miss any error
    smart.errors_as_exceptions(True)

    # The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
    mapping_sheet_column_map = {}
    project_type_deletion_dict = {}
    log_column_map = {}

    report_id = 7437186292311940
    mapping_sheet_id = 3020984954447748
    log_id = 7987758149986180

    mapping_sheet = smart.Sheets.get_sheet(mapping_sheet_id)
    log_sheet = smart.Sheets.get_sheet(log_id)

    for column in log_sheet.columns:
        log_column_map[column.title] = column.id

    for column in mapping_sheet.columns:
        mapping_sheet_column_map[column.title] = column.id

    rowsToDelete = []

    for row in mapping_sheet.rows:
        plan_type = get_cell_by_column_name(row, "Project Plan Type", mapping_sheet_column_map).display_value
        proj_type = get_cell_by_column_name(row, "Project Type", mapping_sheet_column_map).display_value
        project_type_deletion_dict[plan_type] = proj_type

    report = ProjectPlanFactory(smart, report_id, mapping_sheet, mapping_sheet_column_map, project_type_deletion_dict,
                                log_id, log_column_map, logger)

    report.generate_project_plans()

    logger.info("Finished")


