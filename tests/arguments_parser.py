import argparse
import os


def get_cli_arguments():
    parser = argparse.ArgumentParser(description="Universal parser configuration for all WZU utilities")
    parser.add_argument("-l", "--app_url", type=str, help="Application URL")
    parser.add_argument("--tc_image_name", type=str, help="Docker image name")
    parser.add_argument("--flush_tmp_files", type=str, help="Write intermediate temporary files to disk")
    args, unknown = parser.parse_known_args()
    return args


def get_argument(param_name, is_required, err_msg=None):
    args = get_cli_arguments()
    param_value = os.environ.get(param_name)
    if not param_value:
        param_value = getattr(args, param_name.lower(), None)
    if is_required and param_value is None:
        raise ValueError(err_msg)
    return param_value
