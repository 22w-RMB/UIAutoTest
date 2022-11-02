import common
import os


class FilePathHelper:

    @staticmethod
    def get_package_path(package):

        return os.path.dirname(package.__file__)

    @staticmethod
    def get_project_path():

        return os.path.dirname(os.path.dirname(__file__))


if __name__ == '__main__':
    print(FilePathHelper.get_package_path(common))
    print(FilePathHelper.get_project_path())
