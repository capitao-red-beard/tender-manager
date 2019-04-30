
$folder = "\\legros\Data\admin\leo121\revert_to_client\"
$pythonScript = "C:\Users\AV10\IdeaProjects\tender\main.py"
$filter = ".xlsx"


 pushd $folder
 $files = ls
 Foreach ($file in $files) {
    write-host "found file"
    if ($file.extension -eq $filter ) {
    write-host "processing file" 
    $filepath = $folder + $file.name  
    . "C:\Users\AV10\AppData\Local\Programs\Python\Python37-32\python.exe" $pythonScript $filepath} 
    else { 
    write-host "empty" 
    }
 }