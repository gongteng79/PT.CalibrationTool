Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

$ErrorActionPreference = 'Stop'

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Get-Color {
    param([string]$Hex)
    return [System.Drawing.ColorTranslator]::FromHtml($Hex)
}

function New-Canvas {
    param(
        [string]$Path,
        [int]$Width,
        [int]$Height
    )

    $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear([System.Drawing.Color]::White)

    return [pscustomobject]@{
        Path = $Path
        Bitmap = $bitmap
        Graphics = $graphics
    }
}

function Save-Canvas {
    param($Canvas)
    $Canvas.Bitmap.Save($Canvas.Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $Canvas.Graphics.Dispose()
    $Canvas.Bitmap.Dispose()
}

function New-Font {
    param(
        [float]$Size,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )
    return New-Object System.Drawing.Font('Microsoft YaHei UI', $Size, $Style, [System.Drawing.GraphicsUnit]::Pixel)
}

function New-Brush {
    param([string]$Hex)
    return New-Object System.Drawing.SolidBrush (Get-Color $Hex)
}

function New-Pen {
    param(
        [string]$Hex,
        [float]$Width = 2,
        [switch]$Dashed,
        [switch]$Arrow
    )
    $pen = New-Object System.Drawing.Pen (Get-Color $Hex), $Width
    if ($Dashed) {
        $pen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash
    }
    if ($Arrow) {
        $cap = New-Object System.Drawing.Drawing2D.AdjustableArrowCap 5, 6, $true
        $pen.CustomEndCap = $cap
    }
    return $pen
}

function New-RoundedRectPath {
    param(
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [float]$Radius
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $diameter = $Radius * 2
    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
    $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function Draw-TextCentered {
    param(
        $Graphics,
        [string]$Text,
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [float]$FontSize = 24,
        [switch]$Bold,
        [string]$Color = '#1F2937'
    )

    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $font = New-Font -Size $FontSize -Style $style
    $brush = New-Brush $Color
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $format.Trimming = [System.Drawing.StringTrimming]::EllipsisWord

    $rect = [System.Drawing.RectangleF]::new($X, $Y, $Width, $Height)
    $Graphics.DrawString($Text, $font, $brush, $rect, $format)

    $brush.Dispose()
    $font.Dispose()
    $format.Dispose()
}

function Draw-TextLeft {
    param(
        $Graphics,
        [string]$Text,
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [float]$FontSize = 18,
        [switch]$Bold,
        [string]$Color = '#334155'
    )

    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $font = New-Font -Size $FontSize -Style $style
    $brush = New-Brush $Color
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Near
    $format.LineAlignment = [System.Drawing.StringAlignment]::Near

    $rect = [System.Drawing.RectangleF]::new($X, $Y, $Width, $Height)
    $Graphics.DrawString($Text, $font, $brush, $rect, $format)

    $brush.Dispose()
    $font.Dispose()
    $format.Dispose()
}

function Draw-Node {
    param(
        $Graphics,
        [string]$Text,
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [string]$Fill = '#E8F1FB',
        [string]$Border = '#3B82F6',
        [float]$FontSize = 24
    )

    $path = New-RoundedRectPath -X $X -Y $Y -Width $Width -Height $Height -Radius 16
    $fillBrush = New-Brush $Fill
    $borderPen = New-Pen -Hex $Border -Width 2
    $Graphics.FillPath($fillBrush, $path)
    $Graphics.DrawPath($borderPen, $path)
    Draw-TextCentered -Graphics $Graphics -Text $Text -X $X -Y $Y -Width $Width -Height $Height -FontSize $FontSize

    $fillBrush.Dispose()
    $borderPen.Dispose()
    $path.Dispose()
}

function Draw-Container {
    param(
        $Graphics,
        [string]$Title,
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [string]$Fill = '#F8FAFC',
        [string]$Border = '#94A3B8'
    )

    $path = New-RoundedRectPath -X $X -Y $Y -Width $Width -Height $Height -Radius 18
    $fillBrush = New-Brush $Fill
    $borderPen = New-Pen -Hex $Border -Width 2 -Dashed
    $Graphics.FillPath($fillBrush, $path)
    $Graphics.DrawPath($borderPen, $path)
    Draw-TextLeft -Graphics $Graphics -Text $Title -X ($X + 16) -Y ($Y + 10) -Width ($Width - 32) -Height 36 -FontSize 22 -Bold

    $fillBrush.Dispose()
    $borderPen.Dispose()
    $path.Dispose()
}

function Draw-Decision {
    param(
        $Graphics,
        [string]$Text,
        [float]$CenterX,
        [float]$CenterY,
        [float]$Width,
        [float]$Height,
        [string]$Fill = '#FFF4E5',
        [string]$Border = '#F59E0B',
        [float]$FontSize = 22
    )

    $points = @(
        ([System.Drawing.PointF]::new($CenterX, $CenterY - $Height / 2)),
        ([System.Drawing.PointF]::new($CenterX + $Width / 2, $CenterY)),
        ([System.Drawing.PointF]::new($CenterX, $CenterY + $Height / 2)),
        ([System.Drawing.PointF]::new($CenterX - $Width / 2, $CenterY))
    )

    $fillBrush = New-Brush $Fill
    $borderPen = New-Pen -Hex $Border -Width 2
    $Graphics.FillPolygon($fillBrush, $points)
    $Graphics.DrawPolygon($borderPen, $points)
    Draw-TextCentered -Graphics $Graphics -Text $Text -X ($CenterX - $Width / 2 + 8) -Y ($CenterY - $Height / 2 + 8) -Width ($Width - 16) -Height ($Height - 16) -FontSize $FontSize

    $fillBrush.Dispose()
    $borderPen.Dispose()
}

function Draw-ArrowLine {
    param(
        $Graphics,
        [float[]]$Points,
        [string]$Label = '',
        [string]$Color = '#475569'
    )

    $pts = New-Object System.Collections.Generic.List[System.Drawing.PointF]
    for ($i = 0; $i -lt $Points.Length; $i += 2) {
        $pts.Add(([System.Drawing.PointF]::new($Points[$i], $Points[$i + 1])))
    }

    $pen = New-Pen -Hex $Color -Width 3 -Arrow
    for ($i = 0; $i -lt $pts.Count - 2; $i++) {
        $Graphics.DrawLine($pen, $pts[$i], $pts[$i + 1])
    }
    $Graphics.DrawLine($pen, $pts[$pts.Count - 2], $pts[$pts.Count - 1])

    if ($Label) {
        $mid = [int](($pts[0].X + $pts[$pts.Count - 1].X) / 2)
        $midY = [int](($pts[0].Y + $pts[$pts.Count - 1].Y) / 2)
        Draw-TextCentered -Graphics $Graphics -Text $Label -X ($mid - 70) -Y ($midY - 18) -Width 140 -Height 30 -FontSize 18 -Bold -Color '#334155'
    }

    $pen.Dispose()
}

function Draw-Line {
    param(
        $Graphics,
        [float]$X1,
        [float]$Y1,
        [float]$X2,
        [float]$Y2,
        [string]$Color = '#64748B',
        [float]$Width = 2,
        [switch]$Dashed
    )
    $pen = New-Pen -Hex $Color -Width $Width -Dashed:$Dashed
    $Graphics.DrawLine($pen, $X1, $Y1, $X2, $Y2)
    $pen.Dispose()
}

function Draw-Entity {
    param(
        $Graphics,
        [string]$Title,
        [string[]]$Fields,
        [float]$X,
        [float]$Y,
        [float]$Width,
        [string]$HeaderFill = '#E8F1FB',
        [string]$Border = '#3B82F6'
    )

    $height = 54 + ($Fields.Count * 28) + 18
    $path = New-RoundedRectPath -X $X -Y $Y -Width $Width -Height $height -Radius 14
    $fillBrush = New-Brush '#FFFFFF'
    $borderPen = New-Pen -Hex $Border -Width 2
    $Graphics.FillPath($fillBrush, $path)
    $Graphics.DrawPath($borderPen, $path)

    $headerBrush = New-Brush $HeaderFill
    $headerRect = [System.Drawing.RectangleF]::new($X, $Y, $Width, 48)
    $Graphics.FillRectangle($headerBrush, $headerRect)
    Draw-TextCentered -Graphics $Graphics -Text $Title -X $X -Y $Y -Width $Width -Height 48 -FontSize 20 -Bold
    Draw-Line -Graphics $Graphics -X1 $X -Y1 ($Y + 48) -X2 ($X + $Width) -Y2 ($Y + 48) -Color $Border

    $fieldText = ($Fields -join [Environment]::NewLine)
    Draw-TextLeft -Graphics $Graphics -Text $fieldText -X ($X + 12) -Y ($Y + 58) -Width ($Width - 24) -Height ($height - 66) -FontSize 18

    $headerBrush.Dispose()
    $fillBrush.Dispose()
    $borderPen.Dispose()
    $path.Dispose()
}
function Draw-SequenceDiagram {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1600 -Height 920
    $g = $canvas.Graphics

    $actors = @(
        @{ Name = '操作员'; X = 90 },
        @{ Name = 'WPF界面'; X = 300 },
        @{ Name = '流程引擎'; X = 530 },
        @{ Name = '设备适配层'; X = 780 },
        @{ Name = '算法服务'; X = 1040 },
        @{ Name = '仓储'; X = 1280 },
        @{ Name = '报表服务'; X = 1480 }
    )

    foreach ($actor in $actors) {
        Draw-Node -Graphics $g -Text $actor.Name -X ($actor.X - 70) -Y 42 -Width 140 -Height 54 -Fill '#E8F1FB' -Border '#3B82F6' -FontSize 18
        Draw-Line -Graphics $g -X1 $actor.X -Y1 108 -X2 $actor.X -Y2 860 -Color '#94A3B8' -Width 2 -Dashed
    }

    $steps = @(
        @{ From = 90; To = 300; Y = 150; Text = '登录并选择料号/SN' },
        @{ From = 300; To = 530; Y = 220; Text = '创建测试任务' },
        @{ From = 530; To = 780; Y = 290; Text = '执行温箱/压力/采集' },
        @{ From = 780; To = 530; Y = 360; Text = '返回原始采样' },
        @{ From = 530; To = 1040; Y = 430; Text = '数据处理与拟合' },
        @{ From = 1040; To = 530; Y = 500; Text = '返回处理/拟合结果' },
        @{ From = 530; To = 780; Y = 570; Text = '下发标定参数' },
        @{ From = 530; To = 1280; Y = 640; Text = '保存过程与结果' },
        @{ From = 530; To = 1480; Y = 710; Text = '生成Excel/Word' },
        @{ From = 1480; To = 300; Y = 780; Text = '返回导出结果与结论' }
    )

    foreach ($step in $steps) {
        Draw-ArrowLine -Graphics $g -Points @($step.From, $step.Y, $step.To, $step.Y) -Label $step.Text
    }

    Save-Canvas $canvas
}

function Draw-OverviewBusinessFlow {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1600 -Height 860
    $g = $canvas.Graphics

    Draw-Node -Graphics $g -Text '手工选配方 / 录入SN' -X 70 -Y 120 -Width 210 -Height 90 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text '设备自检与读取配置' -X 330 -Y 120 -Width 210 -Height 90
    Draw-Node -Graphics $g -Text '温度设定与稳定等待' -X 590 -Y 120 -Width 210 -Height 90
    Draw-Node -Graphics $g -Text '压力控制与首轮采集' -X 850 -Y 120 -Width 210 -Height 90
    Draw-Node -Graphics $g -Text '数据处理与拟合' -X 1110 -Y 120 -Width 190 -Height 90 -Fill '#F3E8FF' -Border '#8B5CF6'
    Draw-Node -Graphics $g -Text '参数下发' -X 1360 -Y 120 -Width 150 -Height 90 -Fill '#FFF4E5' -Border '#F59E0B'

    Draw-Node -Graphics $g -Text '二轮验证采集' -X 1210 -Y 380 -Width 180 -Height 90
    Draw-Node -Graphics $g -Text '验证处理与判定' -X 930 -Y 380 -Width 180 -Height 90 -Fill '#F3E8FF' -Border '#8B5CF6'
    Draw-Decision -Graphics $g -Text '结果合格？' -CenterX 700 -CenterY 425 -Width 180 -Height 110
    Draw-Node -Graphics $g -Text '归档、导出Excel/Word' -X 430 -Y 380 -Width 210 -Height 90 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text '记录异常并提示人工处理' -X 140 -Y 380 -Width 220 -Height 90 -Fill '#FEE2E2' -Border '#DC2626'

    Draw-ArrowLine -Graphics $g -Points @(280,165,330,165)
    Draw-ArrowLine -Graphics $g -Points @(540,165,590,165)
    Draw-ArrowLine -Graphics $g -Points @(800,165,850,165)
    Draw-ArrowLine -Graphics $g -Points @(1060,165,1110,165)
    Draw-ArrowLine -Graphics $g -Points @(1300,165,1360,165)
    Draw-ArrowLine -Graphics $g -Points @(1435,210,1435,260,1300,260,1300,380)
    Draw-ArrowLine -Graphics $g -Points @(1210,425,1110,425)
    Draw-ArrowLine -Graphics $g -Points @(930,425,790,425)
    Draw-ArrowLine -Graphics $g -Points @(610,425,640,425) -Label '是'
    Draw-ArrowLine -Graphics $g -Points @(610,470,610,560,250,560,250,470) -Label '否'
    Draw-ArrowLine -Graphics $g -Points @(640,425,430,425)

    Save-Canvas $canvas
}

function Draw-LayerArchitecture {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1600 -Height 980
    $g = $canvas.Graphics

    Draw-Container -Graphics $g -Title '表现层 PT.Calibration（WPF + MVVM）' -X 90 -Y 70 -Width 1420 -Height 170 -Fill '#EFF6FF' -Border '#3B82F6'
    Draw-Node -Graphics $g -Text '登录 / Shell / 工位主界面' -X 150 -Y 130 -Width 270 -Height 72
    Draw-Node -Graphics $g -Text '配方管理 / 历史查询 / 报表中心' -X 470 -Y 130 -Width 320 -Height 72
    Draw-Node -Graphics $g -Text '设备配置 / 用户与权限 / 诊断' -X 860 -Y 130 -Width 300 -Height 72
    Draw-Node -Graphics $g -Text 'ViewModel / 导航 / 命令绑定' -X 1210 -Y 130 -Width 250 -Height 72

    Draw-Container -Graphics $g -Title '应用层 PT.Calibration.Application' -X 90 -Y 280 -Width 1420 -Height 180 -Fill '#F8FAFC' -Border '#64748B'
    Draw-Node -Graphics $g -Text 'ITestWorkflowEngine' -X 160 -Y 345 -Width 230 -Height 70 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text 'ITestSessionService / IRecipeService' -X 440 -Y 345 -Width 330 -Height 70 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text 'IReportService / IAuthService / IAuditService' -X 830 -Y 345 -Width 360 -Height 70 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text 'DTO / Command / Query / Policy' -X 1250 -Y 345 -Width 200 -Height 70 -Fill '#E7F7ED' -Border '#16A34A'

    Draw-Container -Graphics $g -Title '领域层 PT.Calibration.Domain' -X 90 -Y 500 -Width 1420 -Height 180 -Fill '#FAF5FF' -Border '#8B5CF6'
    Draw-Node -Graphics $g -Text 'Recipe / TestRun / SampleRecord' -X 160 -Y 565 -Width 300 -Height 70 -Fill '#F3E8FF' -Border '#8B5CF6'
    Draw-Node -Graphics $g -Text 'ProcessResult / FitResult / CalibrationParameter' -X 500 -Y 565 -Width 420 -Height 70 -Fill '#F3E8FF' -Border '#8B5CF6'
    Draw-Node -Graphics $g -Text '设备接口与状态机' -X 970 -Y 565 -Width 220 -Height 70 -Fill '#F3E8FF' -Border '#8B5CF6'
    Draw-Node -Graphics $g -Text '规则、异常、领域事件' -X 1230 -Y 565 -Width 220 -Height 70 -Fill '#F3E8FF' -Border '#8B5CF6'

    Draw-Container -Graphics $g -Title '基础设施层 PT.Calibration.Infrastructure / Simulator' -X 90 -Y 720 -Width 1420 -Height 180 -Fill '#FFF7ED' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text 'SQLite + EF Core 8' -X 160 -Y 785 -Width 240 -Height 70 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text 'Serilog / 模板 / 文件存储' -X 470 -Y 785 -Width 280 -Height 70 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '设备驱动适配器（串口 / TCP / Modbus）' -X 810 -Y 785 -Width 360 -Height 70 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '模拟器 / 报表导出 / 密码哈希' -X 1210 -Y 785 -Width 260 -Height 70 -Fill '#FFF4E5' -Border '#F59E0B'

    Draw-ArrowLine -Graphics $g -Points @(800,240,800,280)
    Draw-ArrowLine -Graphics $g -Points @(800,460,800,500)
    Draw-ArrowLine -Graphics $g -Points @(800,680,800,720)

    Save-Canvas $canvas
}

function Draw-ModuleRelation {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1600 -Height 980
    $g = $canvas.Graphics

    Draw-Node -Graphics $g -Text '工位主界面 / Shell' -X 640 -Y 410 -Width 320 -Height 110 -Fill '#E7F7ED' -Border '#16A34A' -FontSize 28
    Draw-Node -Graphics $g -Text '配方管理' -X 220 -Y 140 -Width 220 -Height 90
    Draw-Node -Graphics $g -Text '设备配置' -X 1160 -Y 140 -Width 220 -Height 90
    Draw-Node -Graphics $g -Text '测试流程引擎' -X 220 -Y 700 -Width 220 -Height 90 -Fill '#F3E8FF' -Border '#8B5CF6'
    Draw-Node -Graphics $g -Text '数据处理与拟合' -X 1160 -Y 700 -Width 220 -Height 90 -Fill '#F3E8FF' -Border '#8B5CF6'
    Draw-Node -Graphics $g -Text '历史查询 / 报表中心' -X 110 -Y 410 -Width 260 -Height 90 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '用户 / 角色 / 审计' -X 1230 -Y 410 -Width 250 -Height 90 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text 'SQLite / 文件存储' -X 650 -Y 140 -Width 300 -Height 90 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '设备适配层 / 模拟器' -X 640 -Y 700 -Width 320 -Height 90 -Fill '#E8F1FB' -Border '#3B82F6'

    Draw-ArrowLine -Graphics $g -Points @(440,185,640,410)
    Draw-ArrowLine -Graphics $g -Points @(1160,185,960,410)
    Draw-ArrowLine -Graphics $g -Points @(440,745,640,520)
    Draw-ArrowLine -Graphics $g -Points @(1160,745,960,520)
    Draw-ArrowLine -Graphics $g -Points @(370,455,640,455)
    Draw-ArrowLine -Graphics $g -Points @(1230,455,960,455)
    Draw-ArrowLine -Graphics $g -Points @(800,230,800,410)
    Draw-ArrowLine -Graphics $g -Points @(800,700,800,520)

    Save-Canvas $canvas
}

function Draw-Deployment {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1600 -Height 920
    $g = $canvas.Graphics

    Draw-Container -Graphics $g -Title '单机工位 PC（Windows 10/11 + .NET 8 Desktop Runtime）' -X 220 -Y 120 -Width 760 -Height 620 -Fill '#F8FAFC' -Border '#64748B'
    Draw-Node -Graphics $g -Text 'PT.Calibration WPF客户端' -X 300 -Y 210 -Width 300 -Height 80 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text 'SQLite 本地数据库' -X 650 -Y 210 -Width 240 -Height 80 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '日志 / 报表 / 模板目录' -X 300 -Y 350 -Width 300 -Height 80 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '设备适配器 + 模拟器' -X 650 -Y 350 -Width 240 -Height 80 -Fill '#E8F1FB' -Border '#3B82F6'
    Draw-Node -Graphics $g -Text '本地账号 / 角色 / 审计' -X 480 -Y 500 -Width 260 -Height 80 -Fill '#F3E8FF' -Border '#8B5CF6'

    Draw-Node -Graphics $g -Text '压力控制器' -X 1120 -Y 150 -Width 200 -Height 74
    Draw-Node -Graphics $g -Text '万用表' -X 1120 -Y 260 -Width 200 -Height 74
    Draw-Node -Graphics $g -Text '高低温试验箱' -X 1120 -Y 370 -Width 200 -Height 74
    Draw-Node -Graphics $g -Text '切换矩阵' -X 1120 -Y 480 -Width 200 -Height 74
    Draw-Node -Graphics $g -Text 'DUT 参数通道' -X 1120 -Y 590 -Width 200 -Height 74
    Draw-Node -Graphics $g -Text '操作员' -X 70 -Y 260 -Width 110 -Height 74 -Fill '#E7F7ED' -Border '#16A34A'

    Draw-ArrowLine -Graphics $g -Points @(180,297,300,250) -Label '界面操作'
    Draw-ArrowLine -Graphics $g -Points @(890,390,1120,187)
    Draw-ArrowLine -Graphics $g -Points @(890,390,1120,297)
    Draw-ArrowLine -Graphics $g -Points @(890,390,1120,407)
    Draw-ArrowLine -Graphics $g -Points @(890,390,1120,517)
    Draw-ArrowLine -Graphics $g -Points @(890,390,1120,627)
    Draw-ArrowLine -Graphics $g -Points @(600,290,650,250)
    Draw-ArrowLine -Graphics $g -Points @(600,390,650,390)
    Draw-ArrowLine -Graphics $g -Points @(740,430,740,500)

    Save-Canvas $canvas
}
function Draw-RolePermission {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1600 -Height 880
    $g = $canvas.Graphics

    Draw-Node -Graphics $g -Text 'Operator' -X 120 -Y 140 -Width 220 -Height 90 -Fill '#E7F7ED' -Border '#16A34A' -FontSize 28
    Draw-Node -Graphics $g -Text 'Engineer' -X 120 -Y 380 -Width 220 -Height 90 -Fill '#E8F1FB' -Border '#3B82F6' -FontSize 28
    Draw-Node -Graphics $g -Text 'Admin' -X 120 -Y 620 -Width 220 -Height 90 -Fill '#F3E8FF' -Border '#8B5CF6' -FontSize 28

    $permX = 520
    $permW = 320
    Draw-Node -Graphics $g -Text '工位执行、查看结果、导出当前报表' -X $permX -Y 110 -Width $permW -Height 86
    Draw-Node -Graphics $g -Text '维护配方、设备配置、模拟器、诊断' -X $permX -Y 255 -Width $permW -Height 86 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '用户管理、权限矩阵、审计查看' -X $permX -Y 400 -Width $permW -Height 86 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '测试历史、异常追溯、报表复打' -X $permX -Y 545 -Width $permW -Height 86
    Draw-Node -Graphics $g -Text '系统参数、模板、归档策略' -X $permX -Y 690 -Width $permW -Height 86 -Fill '#FFF4E5' -Border '#F59E0B'

    Draw-Node -Graphics $g -Text '审计日志' -X 1040 -Y 320 -Width 220 -Height 90 -Fill '#FEE2E2' -Border '#DC2626'
    Draw-Node -Graphics $g -Text '权限矩阵' -X 1040 -Y 520 -Width 220 -Height 90 -Fill '#FEE2E2' -Border '#DC2626'

    Draw-ArrowLine -Graphics $g -Points @(340,185,520,153)
    Draw-ArrowLine -Graphics $g -Points @(340,425,520,298)
    Draw-ArrowLine -Graphics $g -Points @(340,425,520,588)
    Draw-ArrowLine -Graphics $g -Points @(340,665,520,443)
    Draw-ArrowLine -Graphics $g -Points @(340,665,520,733)
    Draw-ArrowLine -Graphics $g -Points @(840,443,1040,365)
    Draw-ArrowLine -Graphics $g -Points @(840,733,1040,565)

    Save-Canvas $canvas
}

function Draw-LoginFlow {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1500 -Height 980
    $g = $canvas.Graphics

    Draw-Node -Graphics $g -Text '启动客户端' -X 600 -Y 60 -Width 260 -Height 80 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text '加载配置 / 日志 / 本地数据库' -X 540 -Y 180 -Width 380 -Height 80
    Draw-Decision -Graphics $g -Text '环境是否正常？' -CenterX 730 -CenterY 340 -Width 240 -Height 120
    Draw-Node -Graphics $g -Text '显示登录页' -X 600 -Y 430 -Width 260 -Height 80
    Draw-Decision -Graphics $g -Text '账号密码校验通过？' -CenterX 730 -CenterY 600 -Width 280 -Height 120 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '加载权限、工位配置、上次会话' -X 540 -Y 730 -Width 380 -Height 80
    Draw-Node -Graphics $g -Text '进入主界面' -X 600 -Y 860 -Width 260 -Height 80 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text '显示错误并禁止进入' -X 150 -Y 300 -Width 280 -Height 80 -Fill '#FEE2E2' -Border '#DC2626'
    Draw-Node -Graphics $g -Text '记录失败次数 / 锁定账号' -X 1100 -Y 560 -Width 260 -Height 80 -Fill '#FEE2E2' -Border '#DC2626'

    Draw-ArrowLine -Graphics $g -Points @(730,140,730,180)
    Draw-ArrowLine -Graphics $g -Points @(730,260,730,280)
    Draw-ArrowLine -Graphics $g -Points @(610,340,430,340) -Label '否'
    Draw-ArrowLine -Graphics $g -Points @(730,400,730,430) -Label '是'
    Draw-ArrowLine -Graphics $g -Points @(730,510,730,540)
    Draw-ArrowLine -Graphics $g -Points @(870,600,1100,600) -Label '否'
    Draw-ArrowLine -Graphics $g -Points @(730,660,730,730) -Label '是'
    Draw-ArrowLine -Graphics $g -Points @(730,810,730,860)

    Save-Canvas $canvas
}

function Draw-StateMachine {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1600 -Height 980
    $g = $canvas.Graphics

    Draw-Node -Graphics $g -Text 'Idle' -X 120 -Y 420 -Width 180 -Height 78 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text 'Configuring' -X 360 -Y 180 -Width 220 -Height 78
    Draw-Node -Graphics $g -Text 'SettingTemperature' -X 670 -Y 120 -Width 250 -Height 78
    Draw-Node -Graphics $g -Text 'ControllingPressure' -X 980 -Y 180 -Width 250 -Height 78
    Draw-Node -Graphics $g -Text 'CollectingBaseline' -X 1240 -Y 340 -Width 240 -Height 78
    Draw-Node -Graphics $g -Text 'ProcessingBaseline' -X 1240 -Y 560 -Width 240 -Height 78
    Draw-Node -Graphics $g -Text 'Fitting' -X 980 -Y 730 -Width 220 -Height 78 -Fill '#F3E8FF' -Border '#8B5CF6'
    Draw-Node -Graphics $g -Text 'DownloadingParameters' -X 670 -Y 790 -Width 270 -Height 78 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text 'CollectingVerification' -X 350 -Y 730 -Width 250 -Height 78
    Draw-Node -Graphics $g -Text 'ProcessingVerification' -X 140 -Y 610 -Width 250 -Height 78
    Draw-Node -Graphics $g -Text 'Completed' -X 120 -Y 160 -Width 180 -Height 78 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text 'Failed / Aborted' -X 1400 -Y 760 -Width 170 -Height 78 -Fill '#FEE2E2' -Border '#DC2626'

    Draw-ArrowLine -Graphics $g -Points @(300,459,360,219)
    Draw-ArrowLine -Graphics $g -Points @(580,219,670,159)
    Draw-ArrowLine -Graphics $g -Points @(920,159,980,219)
    Draw-ArrowLine -Graphics $g -Points @(1230,219,1360,340)
    Draw-ArrowLine -Graphics $g -Points @(1360,418,1360,560)
    Draw-ArrowLine -Graphics $g -Points @(1240,638,1090,730)
    Draw-ArrowLine -Graphics $g -Points @(980,769,940,829)
    Draw-ArrowLine -Graphics $g -Points @(670,829,600,769)
    Draw-ArrowLine -Graphics $g -Points @(350,769,265,688)
    Draw-ArrowLine -Graphics $g -Points @(140,649,210,238)
    Draw-ArrowLine -Graphics $g -Points @(390,649,300,649,300,238) -Label 'OK'
    Draw-ArrowLine -Graphics $g -Points @(1200,769,1400,799) -Label '失败'
    Draw-ArrowLine -Graphics $g -Points @(1360,379,1485,760) -Label '异常'

    Save-Canvas $canvas
}

function Draw-Navigation {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1600 -Height 940
    $g = $canvas.Graphics

    Draw-Node -Graphics $g -Text '登录页' -X 680 -Y 80 -Width 240 -Height 86 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text 'Shell 主框架' -X 650 -Y 240 -Width 300 -Height 96 -Fill '#E8F1FB' -Border '#3B82F6'
    Draw-Node -Graphics $g -Text '工位主界面' -X 160 -Y 520 -Width 220 -Height 84
    Draw-Node -Graphics $g -Text '配方管理' -X 430 -Y 520 -Width 220 -Height 84
    Draw-Node -Graphics $g -Text '设备配置' -X 700 -Y 520 -Width 220 -Height 84
    Draw-Node -Graphics $g -Text '测试历史 / 报表中心' -X 970 -Y 520 -Width 260 -Height 84
    Draw-Node -Graphics $g -Text '用户 / 角色 / 审计' -X 1280 -Y 520 -Width 240 -Height 84
    Draw-Node -Graphics $g -Text '诊断页 / 模拟器面板' -X 670 -Y 710 -Width 280 -Height 84 -Fill '#FFF4E5' -Border '#F59E0B'

    Draw-ArrowLine -Graphics $g -Points @(800,166,800,240)
    Draw-ArrowLine -Graphics $g -Points @(800,336,270,520)
    Draw-ArrowLine -Graphics $g -Points @(800,336,540,520)
    Draw-ArrowLine -Graphics $g -Points @(800,336,810,520)
    Draw-ArrowLine -Graphics $g -Points @(800,336,1100,520)
    Draw-ArrowLine -Graphics $g -Points @(800,336,1400,520)
    Draw-ArrowLine -Graphics $g -Points @(800,336,810,710)

    Save-Canvas $canvas
}

function Draw-EntityRelation {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1800 -Height 1150
    $g = $canvas.Graphics

    Draw-Entity -Graphics $g -Title 'Recipe' -Fields @('RecipeId', 'PartNo', 'Name', 'Version', 'PressureRange', 'WaitSeconds', 'SampleFrequencyMs', 'FitThresholdR2', 'IsEnabled') -X 80 -Y 80 -Width 280
    Draw-Entity -Graphics $g -Title 'DeviceProfile' -Fields @('DeviceProfileId', 'PressureControllerModel', 'MultimeterModel', 'MatrixModel', 'ChamberModel', 'ProtocolType', 'ConnectionProfile') -X 420 -Y 80 -Width 300 -HeaderFill '#FFF4E5' -Border '#F59E0B'
    Draw-Entity -Graphics $g -Title 'TestRun' -Fields @('TestRunId', 'SerialNumber', 'PartNo', 'RecipeId', 'OperatorId', 'StartTime', 'EndTime', 'Status', 'FinalResult') -X 780 -Y 80 -Width 280 -HeaderFill '#E7F7ED' -Border '#16A34A'
    Draw-Entity -Graphics $g -Title 'User' -Fields @('UserId', 'LoginName', 'DisplayName', 'PasswordHash', 'Status') -X 1120 -Y 80 -Width 250 -HeaderFill '#F3E8FF' -Border '#8B5CF6'
    Draw-Entity -Graphics $g -Title 'Role' -Fields @('RoleId', 'RoleCode', 'RoleName', 'Description') -X 1420 -Y 80 -Width 240 -HeaderFill '#F3E8FF' -Border '#8B5CF6'

    Draw-Entity -Graphics $g -Title 'SampleRecord' -Fields @('SampleRecordId', 'TestRunId', 'Stage', 'PointIndex', 'Timestamp', 'SetPressure', 'ActualPressure', 'AnalogValue', 'Switch1', 'Switch2', 'Temperature') -X 240 -Y 480 -Width 320
    Draw-Entity -Graphics $g -Title 'ProcessResult' -Fields @('ProcessResultId', 'TestRunId', 'Stage', 'AvgPressure', 'ActionPoint', 'ResetPoint', 'Hysteresis', 'ErrorPercentFs', 'ValidRatio') -X 640 -Y 480 -Width 320 -HeaderFill '#E8F1FB' -Border '#3B82F6'
    Draw-Entity -Graphics $g -Title 'FitResult' -Fields @('FitResultId', 'TestRunId', 'SlopeA', 'InterceptB', 'R2', 'RetryCount', 'IsAccepted') -X 1040 -Y 480 -Width 280 -HeaderFill '#E8F1FB' -Border '#3B82F6'
    Draw-Entity -Graphics $g -Title 'CalibrationParameter' -Fields @('CalibrationParameterId', 'TestRunId', 'ParameterSet', 'DownloadStatus', 'VerifiedAt') -X 1380 -Y 480 -Width 300 -HeaderFill '#FFF4E5' -Border '#F59E0B'

    Draw-Entity -Graphics $g -Title 'AuditLog' -Fields @('AuditLogId', 'UserId', 'ActionType', 'TargetType', 'TargetId', 'BeforeJson', 'AfterJson', 'CreatedAt') -X 1120 -Y 820 -Width 340 -HeaderFill '#FEE2E2' -Border '#DC2626'

    Draw-Line -Graphics $g -X1 360 -Y1 220 -X2 780 -Y2 220 -Color '#64748B' -Width 2
    Draw-Line -Graphics $g -X1 720 -Y1 220 -X2 780 -Y2 220 -Color '#64748B' -Width 2
    Draw-Line -Graphics $g -X1 1060 -Y1 220 -X2 1120 -Y2 220 -Color '#64748B' -Width 2
    Draw-Line -Graphics $g -X1 1370 -Y1 220 -X2 1420 -Y2 220 -Color '#64748B' -Width 2

    Draw-Line -Graphics $g -X1 920 -Y1 340 -X2 920 -Y2 480 -Color '#64748B' -Width 2
    Draw-Line -Graphics $g -X1 920 -Y1 340 -X2 400 -Y2 480 -Color '#64748B' -Width 2
    Draw-Line -Graphics $g -X1 920 -Y1 340 -X2 800 -Y2 480 -Color '#64748B' -Width 2
    Draw-Line -Graphics $g -X1 920 -Y1 340 -X2 1180 -Y2 480 -Color '#64748B' -Width 2
    Draw-Line -Graphics $g -X1 920 -Y1 340 -X2 1530 -Y2 480 -Color '#64748B' -Width 2
    Draw-Line -Graphics $g -X1 1250 -Y1 220 -X2 1290 -Y2 820 -Color '#64748B' -Width 2

    Draw-TextCentered -Graphics $g -Text '1:N' -X 555 -Y 195 -Width 80 -Height 24 -FontSize 16 -Bold
    Draw-TextCentered -Graphics $g -Text '1:N' -X 865 -Y 380 -Width 80 -Height 24 -FontSize 16 -Bold
    Draw-TextCentered -Graphics $g -Text '1:1' -X 1170 -Y 380 -Width 80 -Height 24 -FontSize 16 -Bold
    Draw-TextCentered -Graphics $g -Text '1:N' -X 1260 -Y 520 -Width 80 -Height 24 -FontSize 16 -Bold

    Save-Canvas $canvas
}

function Draw-ReportFlow {
    param([string]$Path)
    $canvas = New-Canvas -Path $Path -Width 1500 -Height 940
    $g = $canvas.Graphics

    Draw-Node -Graphics $g -Text '测试任务完成' -X 600 -Y 70 -Width 260 -Height 80 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text '汇总 TestRun / SampleRecord / Result' -X 500 -Y 190 -Width 460 -Height 80
    Draw-Node -Graphics $g -Text '装配 Excel 数据记录模型' -X 150 -Y 350 -Width 320 -Height 80
    Draw-Node -Graphics $g -Text '装配 Word 测试报告模型' -X 560 -Y 350 -Width 320 -Height 80
    Draw-Node -Graphics $g -Text '归档文件、写入报告索引' -X 970 -Y 350 -Width 320 -Height 80
    Draw-Decision -Graphics $g -Text '导出是否成功？' -CenterX 730 -CenterY 560 -Width 260 -Height 120 -Fill '#FFF4E5' -Border '#F59E0B'
    Draw-Node -Graphics $g -Text '更新 UI 与历史查询' -X 560 -Y 720 -Width 320 -Height 80 -Fill '#E7F7ED' -Border '#16A34A'
    Draw-Node -Graphics $g -Text '记录异常并提示重试' -X 1000 -Y 710 -Width 300 -Height 80 -Fill '#FEE2E2' -Border '#DC2626'

    Draw-ArrowLine -Graphics $g -Points @(730,150,730,190)
    Draw-ArrowLine -Graphics $g -Points @(730,270,310,350)
    Draw-ArrowLine -Graphics $g -Points @(730,270,720,350)
    Draw-ArrowLine -Graphics $g -Points @(730,270,1130,350)
    Draw-ArrowLine -Graphics $g -Points @(310,430,310,560,600,560)
    Draw-ArrowLine -Graphics $g -Points @(720,430,720,500)
    Draw-ArrowLine -Graphics $g -Points @(1130,430,1130,560,860,560)
    Draw-ArrowLine -Graphics $g -Points @(730,620,730,720) -Label '是'
    Draw-ArrowLine -Graphics $g -Points @(860,560,1000,560,1000,750) -Label '否'

    Save-Canvas $canvas
}

function Get-FileUri {
    param([string]$Path)
    return ('file:///' + ($Path -replace '\\', '/'))
}

function Get-Styles {
    return @"
body { font-family: 'Microsoft YaHei UI'; font-size: 11pt; color: #1f2937; line-height: 1.6; margin: 1.6cm 1.8cm; }
h1 { font-size: 28pt; text-align: center; margin: 0; }
h2 { font-size: 18pt; color: #0f172a; margin-top: 22pt; border-bottom: 2px solid #cbd5e1; padding-bottom: 6px; }
h3 { font-size: 14pt; color: #0f172a; margin-top: 16pt; }
h4 { font-size: 12pt; color: #0f172a; margin-top: 10pt; }
p { margin: 6pt 0; }
ul { margin: 4pt 0 10pt 20pt; }
li { margin: 2pt 0; }
table { width: 100%; border-collapse: collapse; margin: 10pt 0 16pt 0; font-size: 10.5pt; }
th, td { border: 1px solid #94a3b8; padding: 6px 8px; vertical-align: top; }
th { background: #e2e8f0; }
.cover { text-align: center; margin-top: 180pt; }
.cover-subtitle { margin-top: 16pt; font-size: 14pt; color: #475569; }
.cover-meta { margin-top: 70pt; font-size: 12pt; line-height: 2; }
.toc { margin-top: 40pt; }
.figure { text-align: center; margin: 12pt 0 20pt 0; }
.figure img { width: 95%; max-width: 17cm; border: 1px solid #cbd5e1; }
.caption { font-size: 10pt; color: #475569; margin-top: 6pt; }
.note { background: #f8fafc; border-left: 4px solid #3b82f6; padding: 8pt 10pt; margin: 10pt 0; }
.risk { background: #fff7ed; border-left: 4px solid #f59e0b; padding: 8pt 10pt; margin: 10pt 0; }
.break { page-break-after: always; }
code { font-family: Consolas, monospace; font-size: 10pt; background: #f1f5f9; padding: 0 3px; }
"@
}

function New-HtmlDocument {
    param(
        [string]$Title,
        [string]$Body
    )

    $styles = Get-Styles
    return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>$Title</title>
    <style>
$styles
    </style>
</head>
<body>
$Body
</body>
</html>
"@
}
function Convert-HtmlToDocx {
    param(
        [string]$HtmlPath,
        [string]$DocxPath
    )

    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($HtmlPath)
        $doc.SaveAs2($DocxPath, 16)
        $doc.Close()
        $word.Quit()
    } finally {
        if ($doc -ne $null) {
            try { $doc.Close() } catch {}
        }
        if ($word -ne $null) {
            try { $word.Quit() } catch {}
        }
    }
}

