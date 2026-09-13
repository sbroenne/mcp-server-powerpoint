#!/usr/bin/env pwsh

$ErrorActionPreference = 'Stop'
$rootDir = Split-Path -Parent $PSScriptRoot

Add-Type -AssemblyName Microsoft.CodeAnalysis
Add-Type -AssemblyName Microsoft.CodeAnalysis.CSharp

function Get-FunctionScope($node) {
    for ($scope = $node.Parent; $null -ne $scope; $scope = $scope.Parent) {
        if ($scope -is [Microsoft.CodeAnalysis.CSharp.Syntax.BaseMethodDeclarationSyntax] -or
            $scope -is [Microsoft.CodeAnalysis.CSharp.Syntax.AnonymousFunctionExpressionSyntax] -or
            $scope -is [Microsoft.CodeAnalysis.CSharp.Syntax.LocalFunctionStatementSyntax] -or
            $scope -is [Microsoft.CodeAnalysis.CSharp.Syntax.AccessorDeclarationSyntax]) {
            return $scope
        }
    }
    return $node.SyntaxTree.GetRoot()
}

function Get-VariableScope($node) {
    for ($scope = $node.Parent; $null -ne $scope; $scope = $scope.Parent) {
        if ($scope -is [Microsoft.CodeAnalysis.CSharp.Syntax.BlockSyntax] -or
            $scope -is [Microsoft.CodeAnalysis.CSharp.Syntax.ForStatementSyntax] -or
            $scope -is [Microsoft.CodeAnalysis.CSharp.Syntax.UsingStatementSyntax]) {
            return $scope
        }
    }
    return Get-FunctionScope $node
}

function Get-UnwrappedExpression($expression) {
    while ($expression -is [Microsoft.CodeAnalysis.CSharp.Syntax.ParenthesizedExpressionSyntax] -or
        $expression -is [Microsoft.CodeAnalysis.CSharp.Syntax.CastExpressionSyntax] -or
        ($expression -is [Microsoft.CodeAnalysis.CSharp.Syntax.PostfixUnaryExpressionSyntax] -and $expression.OperatorToken.Text -ceq '!')) {
        if ($expression -is [Microsoft.CodeAnalysis.CSharp.Syntax.PostfixUnaryExpressionSyntax]) {
            $expression = $expression.Operand
        } else {
            $expression = $expression.Expression
        }
    }
    return $expression
}

function Test-Acquisition($expression) {
    $expression = Get-UnwrappedExpression $expression
    if ($null -eq $expression) { return $false }
    if ($expression -is [Microsoft.CodeAnalysis.CSharp.Syntax.LiteralExpressionSyntax] -or
        $expression -is [Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax] -or
        $expression -is [Microsoft.CodeAnalysis.CSharp.Syntax.DefaultExpressionSyntax]) { return $false }
    if ($expression -is [Microsoft.CodeAnalysis.CSharp.Syntax.MemberAccessExpressionSyntax]) {
        $receiver = Get-UnwrappedExpression $expression.Expression
        if ($receiver -is [Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax] -and
            $receiver.Identifier.ValueText -cin @('ctx', 'context') -and
            $expression.Name.Identifier.ValueText -cin @('Presentation', 'App')) { return $false }
    }
    return $true
}

function Test-VariableReference($reference, $variable, $locals, $parameters) {
    $scope = Get-VariableScope $variable
    if (-not $scope.Span.Contains($reference.Span)) { return $false }
    foreach ($other in $locals) {
        if ($other.Span.Equals($variable.Span) -or
            $other.Identifier.ValueText -cne $variable.Identifier.ValueText) { continue }
        $otherScope = Get-VariableScope $other
        if ($scope.Span.Contains($otherScope.Span) -and
            -not $scope.Span.Equals($otherScope.Span) -and
            $otherScope.Span.Contains($reference.Span)) { return $false }
    }
    foreach ($parameter in $parameters) {
        if ($parameter.Identifier.ValueText -cne $variable.Identifier.ValueText) { continue }
        $parameterScope = Get-FunctionScope $parameter
        if ($scope.Span.Contains($parameterScope.Span) -and
            $parameterScope.Span.Contains($reference.Span)) { return $false }
    }
    return $true
}

$sourceRoot = Join-Path $rootDir 'src'
if (-not (Test-Path $sourceRoot -PathType Container)) { throw "Source directory not found: $sourceRoot" }
$files = @(Get-ChildItem $sourceRoot -Recurse -Filter '*.cs' -File | Where-Object {
    $_.FullName -notmatch '[/\\](obj|bin)[/\\]' -and $_.Name -notmatch '\.(g|generated|designer)\.cs$'
})
if ($files.Count -eq 0) { throw 'No C# source files found.' }

$findings = 0
$acquisitions = 0
$exemptFiles = 0
foreach ($file in $files) {
    $relativePath = [System.IO.Path]::GetRelativePath($rootDir, $file.FullName)
    if ($relativePath -match '^src[/\\]PowerPointMcp\.ComInterop[/\\]Session[/\\](PresentationBatch|PresentationSession|PresentationSessionRegistry|PresentationShutdownService)\.cs$') {
        $exemptFiles++
        continue
    }

    $tree = [Microsoft.CodeAnalysis.CSharp.CSharpSyntaxTree]::ParseText([System.IO.File]::ReadAllText($file.FullName))
    $parseErrors = @($tree.GetDiagnostics() | Where-Object { $_.Severity -eq [Microsoft.CodeAnalysis.DiagnosticSeverity]::Error })
    if ($parseErrors.Count -gt 0) { throw "Cannot parse ${relativePath}: $($parseErrors[0])" }
    $root = $tree.GetRoot()
    $nodes = @($root.DescendantNodes())
    $releases = @($nodes | Where-Object {
        if ($_ -isnot [Microsoft.CodeAnalysis.CSharp.Syntax.InvocationExpressionSyntax]) { return $false }
        $member = $_.Expression
        if ($member -isnot [Microsoft.CodeAnalysis.CSharp.Syntax.MemberAccessExpressionSyntax]) { return $false }
        $receiverName = ($member.Expression.DescendantTokens() | ForEach-Object { $_.ValueText }) -join ''
        $receiverName -cin @('ComUtilities', 'Sbroenne.PowerPointMcp.ComInterop.ComUtilities',
            'global::Sbroenne.PowerPointMcp.ComInterop.ComUtilities') -and
            $member.Name.Identifier.ValueText -ceq 'Release'
    })
    $assignments = @($nodes | Where-Object { $_ -is [Microsoft.CodeAnalysis.CSharp.Syntax.AssignmentExpressionSyntax] })
    $parameters = @($nodes | Where-Object { $_ -is [Microsoft.CodeAnalysis.CSharp.Syntax.ParameterSyntax] })
    $locals = @($nodes | Where-Object {
        $_ -is [Microsoft.CodeAnalysis.CSharp.Syntax.VariableDeclaratorSyntax] -and
        ($_.Parent.Parent -is [Microsoft.CodeAnalysis.CSharp.Syntax.LocalDeclarationStatementSyntax] -or
            $_.Parent.Parent -is [Microsoft.CodeAnalysis.CSharp.Syntax.ForStatementSyntax] -or
            $_.Parent.Parent -is [Microsoft.CodeAnalysis.CSharp.Syntax.UsingStatementSyntax])
    })
    foreach ($variable in $locals) {
        $type = $variable.Parent.Type
        if ($type -is [Microsoft.CodeAnalysis.CSharp.Syntax.NullableTypeSyntax]) { $type = $type.ElementType }
        if ($type -isnot [Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax] -or $type.Identifier.Text -cne 'dynamic') { continue }
        $name = $variable.Identifier.ValueText
        $function = Get-FunctionScope $variable
        $values = @($variable.Initializer.Value)
        $values += @($assignments | Where-Object {
            $_.Left -is [Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax] -and
            $_.Left.Identifier.ValueText -ceq $name -and
            (Test-VariableReference $_ $variable $locals $parameters)
        } | ForEach-Object { $_.Right })
        if (-not @($values | Where-Object { Test-Acquisition $_ }).Count) { continue }
        $acquisitions++
        $matching = @($releases | Where-Object {
            $release = $_
            (Test-VariableReference $release $variable $locals $parameters) -and
            (Get-FunctionScope $release).Span.Equals($function.Span) -and
            @($release.ArgumentList.Arguments | Where-Object {
                $argument = $_.Expression
                while ($argument -is [Microsoft.CodeAnalysis.CSharp.Syntax.ParenthesizedExpressionSyntax] -or
                    ($argument -is [Microsoft.CodeAnalysis.CSharp.Syntax.PostfixUnaryExpressionSyntax] -and $argument.OperatorToken.Text -ceq '!')) {
                    if ($argument -is [Microsoft.CodeAnalysis.CSharp.Syntax.PostfixUnaryExpressionSyntax]) {
                        $argument = $argument.Operand
                    } else {
                        $argument = $argument.Expression
                    }
                }
                $_.RefKindKeyword.Text -ceq 'ref' -and
                $argument -is [Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax] -and
                $argument.Identifier.ValueText -ceq $name
            }).Count -gt 0
        })
        if ($matching.Count -eq 0) {
            $line = $variable.GetLocation().GetLineSpan().StartLinePosition.Line + 1
            Write-Host "${relativePath}:${line}: '$name' has a dynamic acquisition without a matching release in its owning scope."
            $findings++
        }
    }
}

Write-Host "Scanned $($files.Count) source files; $exemptFiles session ownership files exempt."
Write-Host "Checked $acquisitions dynamic acquisition variables; $findings missing releases."
Write-Host 'Coverage: local dynamic acquisitions only; typed PIA ownership, control flow, and release-in-finally are not verified.'
if ($findings -gt 0) { exit 1 }
Write-Host 'No unmatched acquisitions in the supported patterns.'
exit 0
