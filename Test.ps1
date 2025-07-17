# 设置输出编码为 UTF-8
$OutputEncoding = [System.Text.Encoding]::UTF8
 
# 设置控制台输出编码为 UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
 
# 输出中文
Write-Host "你好，世界！"
 
# 设置变量
$file = ".\iplist.txt"
$lineNumber = 1
$command = " echo '---- CPU ----' ; lscpu | grep -i '^cpu(s' ; echo  '---- MEM  ----'; free -m ; echo '---- DISKS:  ----' ; lsblk | grep -E '^v|^s' ; echo '---- ip: ----' ;hostname -I ;"; 
$password = "123456"
 
# 生成一个基于当前时间戳的随机文件名
$timestamp = Get-Date -Format "yyyyMMddHHmmssfff"
$randomFileName = "file_$timestamp.txt"
Write-Host "随机文件名: $randomFileName"
Get-Content $file  | ForEach-Object { 
 
    # Write-Host "lineNumber: $lineNumber, content: $_"  #命令内容不能输出重定向，故用echo 替代
    echo "lineNumber: $lineNumber, content: $_"
    $lineNumber++
    & plink -l root -pw $password $_ -batch  $command 2>&1
 
} 2>&1 | Tee-Object -FilePath ".\$randomFileName" -Append # Tee-Object 可以用于将标准输出重定向到文件，但默认情况下它不会处理标准错误输出