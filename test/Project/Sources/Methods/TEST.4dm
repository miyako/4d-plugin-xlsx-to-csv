//%attributes = {}
C_TEXT:C284($src)
ARRAY TEXT:C222($sheets;0)
C_TEXT:C284($password)

$src:=System folder:C487(Desktop:K41:16)+"SB_data.xlsx"

  //get xlsx sheets ($src;$sheets;$password)

$dst:=System folder:C487(Desktop:K41:16)+"SB_data"

convert xlsx to csv ($src;$sheets;$password;$dst;"";"";2)