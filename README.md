# ETDs
Workflow for UAlbany ETDs

```
pip install -r requirements.txt
```
1. Compare ETD XML from ProQuest with catalog export
2. Examine embargos in ETD CML
3. Pull list of email contacts

4. Compare Proquest ETDs with Grad School and Catalog embargo records
5. Build Spreadsheet for ingest from Catalog data and ETD XML


## ETD Packages

```python
from packages import ETD

etd = ETD()
etd.load(path/to/dir)
```

#### ETD objects have useful data attributes

```python
print (etd.etd_id)
> lastname-ZbNhwXqECGjEEKgUaxAjWL
print (etd.year)
> 2014
print (etd.supplemental)
> True
if etd.supplemental:
	print (self.supplemental_files)
> /full/path/to/lastname-ZbNhwXqECGjEEKgUaxAjWL/data/lastname_185
```

#### ETD objects also contain useful paths

```python
print (etd.year_dir)
> /full/path/to/ETDs/2014
print (etd.bagDir)
> /full/path/to/ETDs/2014/lastname-ZbNhwXqECGjEEKgUaxAjWL
print (etd.data)
> /full/path/to/ETDs/2014/lastname-ZbNhwXqECGjEEKgUaxAjWL/data
print (etd.xml_file)
> /full/path/to/ETDs/2014/lastname-ZbNhwXqECGjEEKgUaxAjWL/data/lastname_sunyalb_0668A_10534_DATA.xml
print (etd.pdf_file)
> /full/path/to/ETDs/2014/lastname-ZbNhwXqECGjEEKgUaxAjWL/data/lastname_sunyalb_0668A_10534.pdf
```

### etd.bag is a [bagit](https://github.com/LibraryOfCongress/bagit-python) object

bag.info is a dict of bag-info.txt metadata

```python
print (etd.bag.info["Author"])
> Fullname O. Author
print (etd.bag.info["Submitted-Title"])
> My very Longwinded Dissertation
print (etd.bag.info["Bagging-Date"])
> 2023-06-01T14:53:29.785908
print (etd.bag.info["Completion-Date"])
> 2014
if etd.bag.info["Embargo"] == "True":
	print (etd.bag.info["Embargo-Date"])
> 2025-05-09
print (etd.bag.info["ProQuest-Email"])
> my_old_email@albany.edu
print (etd.bag.info["XML-ID"])
> lastname_sunyalb_0668A_10534
print (etd.bag.info["Zip-ID"])
> etdadmin_upload_145229
```

The ETD bag-info.txt metadata should validate via [Bagit-Profiles](https://bagit-profiles.github.io/bagit-profiles-specification/) against [https://archives.albany.edu/static/bagitprofiles/etd-profile-v0.1.json](https://archives.albany.edu/static/bagitprofiles/etd-profile-v0.1.json).

#### Validate it just like [bagit python](https://github.com/LibraryOfCongress/bagit-python)

```python
if etd.bag.is_valid():
    print("yay :)")
else:
    print("boo :(")
> yay :)
```