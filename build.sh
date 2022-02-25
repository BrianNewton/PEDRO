rm -rf dist/ build/
pyinstaller --onedir --add-binary "/Users/briannewton/OneDrive - University of Waterloo/Wetlands lab/Code/.peat-intel/lib/python3.9/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar:./tabula/" PEDRO.py
