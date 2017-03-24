@echo off

echo "begin transform"
time /t

for /R "generated_notes\" %%f in (*.docx) do (
  echo OfficeToPDF.exe %%f generated_notes
  OfficeToPDF.exe %%f generated_notes
)

echo "end transform"
time /t
