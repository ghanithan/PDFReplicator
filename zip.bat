cd output\PDF
mkdir %~dp0\output\PDF_compressed
for /D %%i in (*) do ( 
	cd %%i
	mkdir %~dp0\output\PDF_compressed\%%i
	for /D %%j in (*) do ( 	
	powershell -command "& {& Compress-Archive -Path .\%%j -DestinationPath %~dp0\output\PDF_compressed\%%i\%%j -CompressionLevel Fastest}"
)
	cd ..
)
pause