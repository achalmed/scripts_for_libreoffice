#!/bin/bash

# Convert DOCX and DOC to ODT
find . -type f -name "*.docx" -exec sh -c 'soffice --headless --convert-to odt --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.doc" -exec sh -c 'soffice --headless --convert-to odt --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.dotx" -exec sh -c 'soffice --headless --convert-to ott --outdir "$(dirname "{}")" "{}"' \;

# Convert XLSX and XLS to ODS
find . -type f -name "*.xlsx" -exec sh -c 'soffice --headless --convert-to ods --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.xls" -exec sh -c 'soffice --headless --convert-to ods --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.xlsm" -exec sh -c 'soffice --headless --convert-to ods --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.xltx" -exec sh -c 'soffice --headless --convert-to ots --outdir "$(dirname "{}")" "{}"' \;

# Convert PPTX and PPT to ODP
find . -type f -name "*.pptx" -exec sh -c 'soffice --headless --convert-to odp --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.ppt" -exec sh -c 'soffice --headless --convert-to odp --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.pptm" -exec sh -c 'soffice --headless --convert-to odp --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.ppsx" -exec sh -c 'soffice --headless --convert-to odp --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.pps" -exec sh -c 'soffice --headless --convert-to odp --outdir "$(dirname "{}")" "{}"' \;
find . -type f -name "*.potx" -exec sh -c 'soffice --headless --convert-to otp --outdir "$(dirname "{}")" "{}"' \;



find . -type f -name "*.docx" -exec rm {} \;
find . -type f -name "*.doc" -exec rm {} \;
find . -type f -name "*.dotx" -exec rm {} \;
find . -type f -name "*.xlsx" -exec rm {} \;
find . -type f -name "*.xls" -exec rm {} \;
find . -type f -name "*.xlsm" -exec rm {} \;
find . -type f -name "*.xltx" -exec rm {} \;
find . -type f -name "*.pptx" -exec rm {} \;
find . -type f -name "*.ppt" -exec rm {} \;
find . -type f -name "*.pptm" -exec rm {} \;
find . -type f -name "*.ppsx" -exec rm {} \;
find . -type f -name "*.pps" -exec rm {} \;
find . -type f -name "*.potx" -exec rm {} \;

