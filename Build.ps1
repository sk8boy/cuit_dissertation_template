<#
.SYNOPSIS
    Word 模板打包脚本 - 将 Ribbon XML、VBA 代码等资源打包成 DOTM 文件
.DESCRIPTION
    此脚本将以下组件打包成 Word 模板 (.dotm) 文件:
    - CustomUI (Ribbon XML 和图像)
    - VBA 代码 (从 .bas 文件)
    - 文档模板本身 (.dotx)
.NOTES
    版本: 1.0
    作者: 王铁军
    日期: $(Get-Date -Format "yyyy-MM-dd")
#>

# 参数定义
param (
    [string]$ProjectName = "cuit_dissertation_template",
    [string]$ProjectCNName = "成都信息工程大学硕士学位论文模板",
    [string]$OutputPath = ".\Release",
    [string]$SourcePath = ".\Source",
    [switch]$CleanBeforeBuild
)

# # 设置输出编码为 UTF-8
# $OutputEncoding = [System.Text.Encoding]::UTF8
# # 设置控制台输入输出编码为UTF-8
# [Console]::InputEncoding = [System.Text.Encoding]::UTF8
# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# # 设置PowerShell默认编码为UTF-8
# $PSDefaultParameterValues['*:Encoding'] = 'utf8'

# 创建必要的目录
$null = New-Item -ItemType Directory -Path $OutputPath -Force
$tempDir = Join-Path $env:TEMP "WordTemplateBuild"
$tempDotmPath = Join-Path $tempDir "$ProjectName.dotm"

# 清理之前的构建
if ($CleanBeforeBuild) {
    Write-Host "清理之前的构建文件..."
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    # Remove-Item -Path (Join-Path $OutputPath "$ProjectName.dotm") -Force -ErrorAction SilentlyContinue
}
# 创建临时目录
$null = New-Item -ItemType Directory -Path $tempDir -Force

# 1. 处理主模板文件
$dotmPath = Join-Path $OutputPath "$ProjectName.dotm"
if (Test-Path $dotmPath) {
    # 复制 .dotm 到临时目录并重命名为 .zip
    $tempZipPath = Join-Path $tempDir "$ProjectName.zip"
    Copy-Item $dotmPath $tempZipPath
    # 解压 ZIP 文件
    Expand-Archive -Path $tempZipPath -DestinationPath $tempDotmPath -Force
    # 删除 ZIP 文件
    Remove-Item -Path $tempZipPath -Recurse -Force -ErrorAction Stop
}
else {
    Write-Error "找不到主模板文件: $dotmPath"
    exit 1
}

# 2. 添加 CustomUI (Ribbon XML)
$customUIPath = Join-Path $SourcePath "CustomUI"
if (Test-Path $customUIPath) {
    Write-Host "添加 CustomUI 组件..."
    
    $customUITargetPath = Join-Path $tempDotmPath "customUI"
    # 删除原dotm文件中自带的 customUI 目录
    Remove-Item -Path $customUITargetPath -Recurse -Force -ErrorAction Stop

    # 创建一个空的 customUI 目录
    $null = New-Item -ItemType Directory -Path $customUITargetPath -Force
    # 复制 Source \ CustomUI 目录下所有文件到空的 customUI 目录
    Copy-Item -Path "$customUIPath\*" -Destination $customUITargetPath -Recurse -Force
    
    # # 清理解压的文件
    # Get-ChildItem -Path $tempDotmPath -Exclude "*.zip" | Remove-Item -Recurse -Force
}

# 3. 添加 VBA 项目

# 初始化 version.txt
$versionFile = ".\Version.txt"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
if (-not (Test-Path $versionFile)) {
    # Set-Content -Path $versionFile -Value "1.0.0"
    [System.IO.File]::WriteAllText($versionFile, "1.0.0", $utf8NoBom)
    $newVersion = 1.0.0
}
else {
    $version = (Get-Content $versionFile).Trim()
    # 拆分版本号（Major.Minor.Patch.Build）
    $versionParts = $version.Split('.')
    $buildNumber = [int]$versionParts[2] + 1
    $newVersion = "$($versionParts[0]).$($versionParts[1]).$buildNumber"

    # 更新 version.txt
    # Set-Content -Path $versionFile -Value $newVersion
    [System.IO.File]::WriteAllText($versionFile, $newVersion, $utf8NoBom)
}

$vbaPath = Join-Path $SourcePath "VBAProject"
if (Test-Path $vbaPath) {
    Write-Host "添加 VBA 组件..."

    # 3.1 拼接 Python 脚本路径（与 ps1 同目录）
    $pyScript = Join-Path $PSScriptRoot "makeVBAProjectFile.py"

    if (-not (Test-Path $pyScript)) {
        Write-Error "未找到配套 Python 脚本：$pyScript"
        exit 1
    }

    # 3.2 调用 Python 提取 vbaProject.bin
    $vbaProjectFile = Join-Path $vbaPath "vbaProject.bin"
    try {
        python $pyScript
        if (Test-Path $vbaProjectFile) {
            Write-Host "提取成功 -> $vbaProjectFile"
        }
        else {
            Write-Error "提取失败，vbaProject.bin 未生成"
            exit 1
        }
    }
    catch {
        Write-Error "Python 执行出错：$_"
        eixt 1
    }

   
    # 这里需要调用 VBA 编译器或使用现有的 vbaProject.bin 文件
    # 这是一个简化版本，假设已经有编译好的 vbaProject.bin
    if (Test-Path $vbaProjectFile) {
        # 目标文件会被存储到 dotm 临时目录下 word/vbaProject.bin
        $vbaProjectPath = Join-Path $tempDotmPath "word" | Join-Path -ChildPath "vbaProject.bin"
        Copy-Item -Path (Join-Path $vbaPath "vbaProject.bin") -Destination $vbaProjectPath -Force
    }
    else {
        Write-Warning "未找到预编译的 vbaProject.bin 文件。VBA 代码将不会包含在模板中"
    }
}

# 4. 重命名为 .dotm 并移动到输出目录
# 重新压缩为 ZIP
Compress-Archive -Path "$tempDotmPath\*" -DestinationPath $tempZipPath -CompressionLevel Optimal -Force
Write-Host "创建最终 DOTM 文件..."
$releaseFileName = "$($ProjectCNName)v$($newVersion).dotm"
Rename-Item -Path $tempZipPath -NewName $releaseFileName -Force
Move-Item -Path (Join-Path $tempDir $releaseFileName) -Destination $OutputPath -Force

Write-Host "构建成功完成! DOTM 文件已保存到: " $(Join-Path $OutputPath $releaseFileName)