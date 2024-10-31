# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Lambda Robot.xlam\*\* contains definitions for:

[23 Robot Commands](#command-definitions)<BR>[13 Robot Parameters](#parameter-definitions)<BR>

<BR>

## Available Robot Commands

[LAMBDA](#lambda) | [LET](#let) | [Paste](#paste)

### LAMBDA

| Name | Description |
| --- | --- |
| [Audit Formula Steps](#audit-formula-steps) | Generate steps for each of the let variable calculation and place all LAMBDA parameters and mark them as input cell. |
| [Cancel Lambda Edit](#cancel-lambda-edit) | Cancel any edits to Lambda definition in active cell and revert back to custom Lambda call. |
| [Cycle LET Steps](#cycle-let-steps) | Cyclically change last steps of the let so that we can see different steps calculated value. |
| [Edit Lambda](#edit-lambda) | Converts a custom Lambda function to it's definition for editing. |
| [Generate Formula Dependency Info](#generate-formula-dependency-info) | Generates a table of formula dependency info for the active cell. The table is placed in first unused space to the right. |
| [Generate Lambda Dependency List](#generate-lambda-dependency-list) | Generate list of Lambdas the ActiveCell formula is dependent on. |
| [Generate Lambda Statement](#generate-lambda-statement) | Generate a Lambda statement based on formula precedents for the active cell and replace ActiveCell formula with generated one. |
| [Generate Multi Column Lookup Lambda](#generate-multi-column-lookup-lambda) | Generate multi column lookup lambda from ActiveCell. It will return nth column value by using n\-1 parameter and filter by them. |
| [Include Lambda Dependencies](#include-lambda-dependencies) | Create Let step for all necessary Lambdas with their definition so that formula became independent of any Lambdas. |
| [Lambda Properties](#lambda-properties) | Edit Lambda properties like Name, Author, Metadata, etc. |
| [LAMBDA To LET](#lambda-to-let) | Convert ActiveCell Lambda formula to LET formula. |
| [LET To LAMBDA](#let-to-lambda) | Convert ActiveCell LET formula to LAMBDA formula. |
| [Mark As Input Cells](#mark-as-input-cells) | Change selected cells background and font color to mark as input cells. |
| [Mark Lambda As LET Step](#mark-lambda-as-let-step) | Create LETStep\_FX and LETStepRef\_FX named range for activecell formula so that we can use them for further calculation for generating lambda statement. |
| [Paste Lambda Statement](#paste-lambda-statement) | Generate a Lambda statement in the active cell based on formula present in copied cell. |
| [Remove Lambda](#remove-lambda) | Removes the defined name for the Lambda in active cell and reverts back to Lambda definition. |
| [Remove Last LET Step](#remove-last-let-step) | Remove last LET Step from ActiveCell formula. if ActiveCell formula \=LET(X,1,Y,2,Add,X+Y,Add) then output will be \=LET(X,1,Y,2,Y). |
| [Remove Unused Lambdas](#remove-unused-lambdas) | Removes all lambdas from the name manager that are not currently being used in the active workbook. |
| [Reset Cycle LET Steps](#reset-cycle-let-steps) | Remove our special identifier if the initial formula last step was calculation or just refer the final steps. |
| [Save Lambda](#save-lambda) | Saves the Lambda definition in the active cell as a named range. |
| [Save Lambda As](#save-lambda-as) | Saves the Lambda definition in the active cell as a new named range specified by user. |

### LET

| Name | Description |
| --- | --- |
| [Audit Formula Steps](#audit-formula-steps) | Generate steps for each of the let variable calculation and place all LAMBDA parameters and mark them as input cell. |
| [Cycle LET Steps](#cycle-let-steps) | Cyclically change last steps of the let so that we can see different steps calculated value. |
| [Generate Let Statement](#generate-let-statement) | Generate a Let statement based on formula precedents for the active cell and replace ActiveCell formula with generated LET. |
| [LAMBDA To LET](#lambda-to-let) | Convert ActiveCell Lambda formula to LET formula. |
| [LET To LAMBDA](#let-to-lambda) | Convert ActiveCell LET formula to LAMBDA formula. |
| [Mark Lambda As LET Step](#mark-lambda-as-let-step) | Create LETStep\_FX and LETStepRef\_FX named range for activecell formula so that we can use them for further calculation for generating lambda statement. |
| [Remove Last LET Step](#remove-last-let-step) | Remove last LET Step from ActiveCell formula. if ActiveCell formula \=LET(X,1,Y,2,Add,X+Y,Add) then output will be \=LET(X,1,Y,2,Y). |
| [Reset Cycle LET Steps](#reset-cycle-let-steps) | Remove our special identifier if the initial formula last step was calculation or just refer the final steps. |

### Paste

| Name | Description |
| --- | --- |
| [Paste Combine Arrays](#paste-combine-arrays) | Paste a dynamic array referencing the copied cells by combining any dynamic arrays automatically. |
| [Paste Lambda Statement](#paste-lambda-statement) | Generate a Lambda statement in the active cell based on formula present in copied cell. |

<BR>

## Available Robot Parameters

| Name | Description |
| --- | --- |
| [FormulaFormat\_AddPrefixOnParam](#formulaformat_addprefixonparam) | Add Let Var Prefix in the parameter as well or not. |
| [FormulaFormat\_BoMode](#formulaformat_bomode) | Set formula formatting option as Bo does. |
| [FormulaFormat\_IndentChar](#formulaformat_indentchar) | Indent Char option for Format Formula. |
| [FormulaFormat\_IndentSize](#formulaformat_indentsize) | Indent Size option for Format Formula. |
| [FormulaFormat\_LambdaParamStyle](#formulaformat_lambdaparamstyle) | Lambda Parameter naming convention. Allowed values are "Pascal", "Camel" or "Snake\_Case". |
| [FormulaFormat\_LetVarPrefix](#formulaformat_letvarprefix) | What prefix to use for each let var. e.g. \_ or var. |
| [FormulaFormat\_Multiline](#formulaformat_multiline) | Multiline option for Format Formula. |
| [FormulaFormat\_OnlyWrapFunctionAfterNChars](#formulaformat_onlywrapfunctionafternchars) | OnlyWrapFunctionAfterNChars option for Format Formula. |
| [FormulaFormat\_SpacesAfterArgumentSeparators](#formulaformat_spacesafterargumentseparators) | SpacesAfterArgumentSeparators option for Format Formula. |
| [FormulaFormat\_SpacesAfterArrayColumnSeparators](#formulaformat_spacesafterarraycolumnseparators) | SpacesAfterArrayColumnSeparators option for Format Formula. |
| [FormulaFormat\_SpacesAfterArrayRowSeparators](#formulaformat_spacesafterarrayrowseparators) | SpacesAfterArrayRowSeparators option for Format Formula. |
| [FormulaFormat\_SpacesAroundInfixOperators](#formulaformat_spacesaroundinfixoperators) | SpacesAroundInfixOperators option for Format Formula. |
| [FormulaFormat\_VariableStyle](#formulaformat_variablestyle) | Let Variable or Parameter naming convention. Allowed values are "Pascal", "Camel" or "Snake\_Case". |

<BR>

## Command Definitions

<BR>

### Audit Formula Steps

*Generate steps for each of the let variable calculation and place all LAMBDA parameters and mark them as input cell.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA` `#LET`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAuditLambdaSteps.GenerateLambdaSteps](./VBA/modAuditLambdaSteps.bas#L15)([[ActiveCell]],[[NewTableTargetToRight]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>afs</code> |

[^Top](#oa-robot-definitions)

<BR>

### Cancel Lambda Edit

*Cancel any edits to Lambda definition in active cell and revert back to custom Lambda call.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaEditor.CancelEditLambda](./VBA/modLambdaEditor.bas#L610)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell AND ExcelActiveCellHasComment |
| Launch Codes | <code>cle</code> |

[^Top](#oa-robot-definitions)

<BR>

### Cycle LET Steps

*Cyclically change last steps of the let so that we can see different steps calculated value.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA` `#LET`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLETStepManager.CycleLETSteps](./VBA/modLETStepManager.bas#L33)([[ActiveCell]],[[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula OR ExcelActiveCellIsSpillParent |
| Launch Codes | <code>cls</code> |

[^Top](#oa-robot-definitions)

<BR>

### Edit Lambda

*Converts a custom Lambda function to it's definition for editing.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaEditor.EditLambda](./VBA/modLambdaEditor.bas#L15)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>el</code> |

[^Top](#oa-robot-definitions)

<BR>

### Generate Formula Dependency Info

*Generates a table of formula dependency info for the active cell. The table is placed in first unused space to the right.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaBuilder.GenerateDependencyInfo](./VBA/modLambdaBuilder.bas#L18)([[ActiveCell]],[[NewTableTargetToRight]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Generate Lambda Dependency List

*Generate list of Lambdas the ActiveCell formula is dependent on.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDependencyFormulaReplacer.GenerateLambdaFormulaDependency](./VBA/modDependencyFormulaReplacer.bas#L58)([[ActiveCell]],[[NewTableTargetToRight]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Generate Lambda Statement

*Generate a Lambda statement based on formula precedents for the active cell and replace ActiveCell formula with generated one.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaBuilder.GenerateLambdaStatement](./VBA/modLambdaBuilder.bas#L115)([[ActiveCell]],[[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>gl</code> |

[^Top](#oa-robot-definitions)

<BR>

### Generate Let Statement

*Generate a Let statement based on formula precedents for the active cell and replace ActiveCell formula with generated LET.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LET`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaBuilder.GenerateLetStatement](./VBA/modLambdaBuilder.bas#L59)([[ActiveCell]],[[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>let</code> |

[^Top](#oa-robot-definitions)

<BR>

### Generate Multi Column Lookup Lambda

*Generate multi column lookup lambda from ActiveCell. It will return nth column value by using n\-1 parameter and filter by them.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[MultiColumnLookupLambda.GenerateMultiColumnLookUpLambda](./VBA/MultiColumnLookupLambda.bas#L14)([[ActiveCell]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Include Lambda Dependencies

*Create Let step for all necessary Lambdas with their definition so that formula became independent of any Lambdas.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDependencyFormulaReplacer.IncludeLambdaDependencies](./VBA/modDependencyFormulaReplacer.bas#L11)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Lambda Properties

*Edit Lambda properties like Name, Author, Metadata, etc.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMetadataEditor.EditLambdaProperties](./VBA/modMetadataEditor.bas#L14)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### LAMBDA To LET

*Convert ActiveCell Lambda formula to LET formula.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA` `#LET`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAuditLambdaSteps.LambdaToLet](./VBA/modAuditLambdaSteps.bas#L48)([[ActiveCell]],[[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### LET To LAMBDA

*Convert ActiveCell LET formula to LAMBDA formula.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA` `#LET`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaBuilder.LetToLambda](./VBA/modLambdaBuilder.bas#L284)([[ActiveCell]],[[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Mark As Input Cells

*Change selected cells background and font color to mark as input cells.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaBuilder.MarkAsInputCells](./VBA/modLambdaBuilder.bas#L590)([[Selection]],False)</code> |
| Launch Codes | <code>i</code> |

[^Top](#oa-robot-definitions)

<BR>

### Mark Lambda As LET Step

*Create LETStep\_FX and LETStepRef\_FX named range for activecell formula so that we can use them for further calculation for generating lambda statement.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA` `#LET`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaBuilder.MarkLambdaAsLETStep](./VBA/modLambdaBuilder.bas#L607)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula OR ExcelActiveCellIsSpillParent |

[^Top](#oa-robot-definitions)

<BR>

### Paste Combine Arrays

*Paste a dynamic array referencing the copied cells by combining any dynamic arrays automatically.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCombineArray.PasteCombineArrays](./VBA/modCombineArray.bas#L11)([[Clipboard]],[[ActiveCell]])</code> |
| User Context Filter | ClipboardHasExcelData AND ExcelActiveCellIsEmpty AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>pca</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Lambda Statement

*Generate a Lambda statement in the active cell based on formula present in copied cell.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#Paste` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaBuilder.GenerateLambdaStatement](./VBA/modLambdaBuilder.bas#L115)([[Clipboard]],[[ActiveCell]])</code> |
| User Context Filter | ClipboardHasExcelData AND ExcelCopiedRangeIsNotEmpty AND ExcelCopiedRangeIsSingleCell |
| Launch Codes | <code>pl</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Lambda

*Removes the defined name for the Lambda in active cell and reverts back to Lambda definition.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaEditor.RemoveLambda](./VBA/modLambdaEditor.bas#L567)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Remove Last LET Step

*Remove last LET Step from ActiveCell formula. if ActiveCell formula \=LET(X,1,Y,2,Add,X+Y,Add) then output will be \=LET(X,1,Y,2,Y).*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA` `#LET`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLETStepManager.RemoveLastLETStep](./VBA/modLETStepManager.bas#L23)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Remove Unused Lambdas

*Removes all lambdas from the name manager that are not currently being used in the active workbook.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modUnUsedLambdas.RemoveUnusedLambdas](./VBA/modUnUsedLambdas.bas#L9)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Reset Cycle LET Steps

*Remove our special identifier if the initial formula last step was calculation or just refer the final steps.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA` `#LET`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLETStepManager.CycleLETSteps](./VBA/modLETStepManager.bas#L33)([[ActiveCell]],[[ActiveCell]],True)</code> |
| User Context Filter | ExcelActiveCellContainsFormula OR ExcelActiveCellIsSpillParent |

[^Top](#oa-robot-definitions)

<BR>

### Save Lambda

*Saves the Lambda definition in the active cell as a named range.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaEditor.SaveLambda](./VBA/modLambdaEditor.bas#L173)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>sl</code> |

[^Top](#oa-robot-definitions)

<BR>

### Save Lambda As

*Saves the Lambda definition in the active cell as a new named range specified by user.*

<sup>`@Lambda Robot.xlam` `!VBA Macro Command` `#LAMBDA`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdaEditor.SaveLambdaAs](./VBA/modLambdaEditor.bas#L462)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>sla</code> |

[^Top](#oa-robot-definitions)

<BR>

## Parameter Definitions

<BR>

### FormulaFormat\_AddPrefixOnParam

*Add Let Var Prefix in the parameter as well or not.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>false</code> |
| Data Type | Boolean |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_BoMode

*Set formula formatting option as Bo does.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>false</code> |
| Data Type | Boolean |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_IndentChar

*Indent Char option for Format Formula.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code> </code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_IndentSize

*Indent Size option for Format Formula.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>3</code> |
| Data Type | Integer |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_LambdaParamStyle

*Lambda Parameter naming convention. Allowed values are "Pascal", "Camel" or "Snake\_Case".*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>Snake\_Case</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_LetVarPrefix

*What prefix to use for each let var. e.g. \_ or var.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>\_</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_Multiline

*Multiline option for Format Formula.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>true</code> |
| Data Type | Boolean |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_OnlyWrapFunctionAfterNChars

*OnlyWrapFunctionAfterNChars option for Format Formula.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>80</code> |
| Data Type | Integer |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_SpacesAfterArgumentSeparators

*SpacesAfterArgumentSeparators option for Format Formula.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>true</code> |
| Data Type | Boolean |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_SpacesAfterArrayColumnSeparators

*SpacesAfterArrayColumnSeparators option for Format Formula.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>true</code> |
| Data Type | Boolean |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_SpacesAfterArrayRowSeparators

*SpacesAfterArrayRowSeparators option for Format Formula.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>true</code> |
| Data Type | Boolean |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_SpacesAroundInfixOperators

*SpacesAroundInfixOperators option for Format Formula.*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>true</code> |
| Data Type | Boolean |

[^Top](#oa-robot-definitions)

<BR>

### FormulaFormat\_VariableStyle

*Let Variable or Parameter naming convention. Allowed values are "Pascal", "Camel" or "Snake\_Case".*

<sup>`@Lambda Robot.xlam` `!Default Parameter` </sup>

| Property | Value |
| --- | --- |
| Default Value | <code>Pascal</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)
