import sys
import pathlib
import shutil
import ruamel.std.zipfile as zipfile
from docx import Document
from os import rmdir, removedirs
from os.path import join, expanduser
from bs4 import BeautifulSoup


mpath = str(pathlib.Path(__file__).parent.absolute())
archive = ""

_image_ = ""
image_path = ""

doc_path = []
rel_path = []

docxx_name = ""

file_xml = []
file_rls = []

# needed files from .docx
__xml_rels__ = ["document.xml", "document.xml.rels", "numbering.xml"]


def merge_docx_file():
    for image in archive.filelist:
        # Get the images (.png, .jpg, etc) and set the path of its file, and
        # move the file to current folder
        if image.filename.startswith('word/media/'):
            # get image name
            _image_ = image.filename.split("/")[-1]
            # set image path
            image_path = mpath + "\\archive\\" + image.filename
            # extract image
            archive.extract(image.filename, path=mpath + "./archive/")

    # move image
    pathlib.Path(image_path).rename(image_path.replace("word/media/", ""))
    # delete folder word
    shutil.rmtree(mpath + "\\archive\word")

    # create docx file
    document = Document()
    document.add_heading('Document Title', 0)
    document.save(docxx_name)

    image_path = mpath + "\\archive\\" + _image_
    with zipfile.ZipFile(image_path) as content:
        try:
            # extract all files from .docx to content directory
            content.extractall(path="./archive/content/")

            for file in content.filelist:
                # Get the file of *.xml from word/ folder and set
                # the path of its file, and move the file to
                # current folder
                if file.filename.startswith("word/") and file.filename.endswith('.xml'):
                    _xml_ = file.filename.split("/")[-1]

                    if _xml_ != "theme1.xml" and _xml_ in __xml_rels__:
                        file_xml.append(_xml_)
                        doc_path.append(
                            mpath + "\\archive\content\\" + file.filename)

                # Get the file of *.rels from word/_rel folder and set
                # the path of its file, and move the file to
                # current folder
                if file.filename.startswith("word/") and file.filename.endswith('.rels'):
                    _rels_ = file.filename.split("/")[-1]

                    if _rels_ in __xml_rels__:
                        file_rls.append(_rels_)
                        rel_path.append(
                            mpath + "\\archive\content\\" + file.filename)

            # create directory xml & rels file
            pathlib.Path(
                mpath + "/archive/xml_rels/").mkdir(parents=True, exist_ok=True)

            # move *.xml to ./archive/xml_rels
            for xml in doc_path:
                pathlib.Path(xml).rename(
                    xml.replace("content\word/", "xml_rels\\"))

            # move *.xml to ./archive/xml_rels
            for rels in rel_path:
                pathlib.Path(rels).rename(rels.replace(
                    "content\word/_rels/", "xml_rels\\"))

            # remove content directory
            while True:
                shutil.rmtree(mpath + "\\archive\content")
                break

        except Exception as ex:
            print("Error caution at : [\n", ex, "\n]")

    # delete .*.xml* in file.docx file, so
    # when the file.docx is open and append the file
    # '.*.xml*' in './archive/.*.xml*' to
    # avoid the duplicated file error when write the
    # new file with the './archive/.*.xml*'
    for file in __xml_rels__:
        zipfile.delete_from_zip_file(docxx_name, pattern=f'.*{file}')

    with zipfile.ZipFile(docxx_name, 'a') as content:
        try:
            for file in __xml_rels__:
                _file_ = f'./archive/xml_rels/{file}'
                with open(_file_, 'r', encoding='utf-8') as f:
                    data=f.read()

                    # read xml content
                    xml_content=BeautifulSoup(data, "xml")

                    if file.endswith(".rels"):
                        # write the rels file
                        content.writestr(f"word/_rels/{file}",
                                         bytes(str(xml_content), encoding='utf-8'))
                    else:
                        # write the xml file
                        content.writestr(f"word/{file}",
                                         bytes(str(xml_content), encoding='utf-8'))

                    f.close()

            while True:
                shutil.rmtree(mpath + "\\archive")
                print("File has been merged.")
                print("Completed.")
                break

        except Exception as ex:
            print(f"Error caution when opening {docxx_name}: [\n", ex, "\n]")


def main():
    """ python run.py -f file.docx -s result.docx
    """
    global archive, docxx_name

    args=sys.argv[1:]
    file1=args[1]
    save2=args[3]

    archive=zipfile.ZipFile(mpath + f'./docx/{file1}.docx')
    docxx_name=save2 + ".docx"

    merge_docx_file()


if __name__ == "__main__":
    main()
