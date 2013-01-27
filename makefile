all: setup.py webprintfolder.pyw
	python setup.py py2exe
	echo 'Build complete. Running Application...'
	cp *.py ../src/
	cp *.pyw ../src/
	cp makefile ../src/makefile
	../dist/webprintfolder.exe

clean:
	rm ../build/* ../dist/*