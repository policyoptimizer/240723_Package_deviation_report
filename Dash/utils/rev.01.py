import dataiku

def get_managed_folder_list():
    """
    Get the list of managed folders in the current project

    :return: A list of (id, name)
    """
    project = dataiku.api_client().get_default_project()
    managed_folders = project.list_managed_folders()
    ids_and_names = [(mf.get('id', ''), mf.get('name', ''))
                     for mf in managed_folders]
    return ids_and_names

def get_files_in_folder(folder_id):
    """
    Get the list of files in a managed folder

    :param id: Id of the managed folder

    :return: A list of files in the managed folder
    """
    mf = dataiku.Folder(folder_id)
    files = mf.list_paths_in_partition()
    return files

