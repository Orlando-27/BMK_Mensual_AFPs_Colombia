# Fórmulas clave — Benchmark Mensual Optimizado

## Fwd Industria fila 23 (AK..BA) - fwd compra/venta y nominal  [Fwd Industria]
- `AN23`: `=IF(AL23="COP","VENTA",IF(AM23="COP","COMPRA",IF(AND(AL23<>"COP",AM23="USD"),"COMPRA",IF(AND(AM23<>"COP",AL23="USD"),"VENTA","REVISAR"))))`
- `AQ23`: `=IF(AN23="compra",AO23,AP23)`
- `AV23`: `=IF(OR(IF(AN23="VENTA",AM23&AL23,AL23&AM23)="USDCOP",IF(AN23="VENTA",AM23&AL23,AL23&AM23)="EURUSD"),IF(AN23="VENTA",AM23&AL23,AL23&AM23),IF(AN23="VENTA",AL23&AM23,AM23&AL23))`
- `AW23`: `=IF(OR(AV23="USDJPY",AV23="USDMXN",AV23="USDBRL",AV23="USDCAD"),1/AR23,AR23)`
- `AX23`: `=IF(OR(AY23="USDCAD"),1/AR23,AR23)`
- `AY23`: `=IF(OR(OR(OR(IF(AN23="VENTA",AM23&AL23,AL23&AM23)="USDCOP",IF(AN23="VENTA",AM23&AL23,AL23&AM23)="EURUSD"),IF(AN23="VENTA",AM23&AL23,AL23&AM23)="GBPUSD"),IF(AN23="VENTA",AM23&AL23,AL23&AM23)="AUDUSD"),IF(AN23="VENTA",AM23&AL23,AL23&AM23),IF(AN23="VENTA",AL23&AM23,AM23&AL23))`

## Swaps fila 2 (BX..CU) - export a Calculadora Swap  [Swaps]
- `BX2`: `=+IF(CI2=CQ2,IF(CJ2="COP",IF(MID(CF2,1,8)="IBR o/n+","IRS IBR+PB","IRS IBRCOP"),IF(AND(CJ2="USD",OR(CE2="SOF1",CM2="SOF1")),"IRS SOFUSD","IRS LIBORUSD")),IF(AND(CJ2="COP",CR2="USD",CN2="LIBOR",LEFT(CF2,3)="IBR"),"CCS IBRLIB",(IF(OR(AND(CJ2="COP",CR2="USD"),AND(CR2="COP",CJ2="USD")),"CCS COPUSD",IF(OR(AND(CJ2="COP",CR2="UVR"),AND(CR2="COP",CJ2="UVR")),"CCS COPUVR",IF(OR(AND(CJ2="COP",CR2="EUR"),AND(CR2="COP",CJ2="EUR")),"CCS COPEUR",IF(OR(AND(CJ2="UVR",CR2="EUR"),AND(CR2="UVR",CJ2="EUR")),"CCS UVREUR",IF(OR(AND(CJ2="UVR",CR2="USD"),AND(CR2="UVR",CJ2="USD")),"CCS UVRUSD"))))))))`
- `BY2`: `=+IF(BZ2="OLD MUTUAL","SK"&L2,IF(BZ2="COLFONDOS","Col"&L2,L2))`
- `BZ2`: `=+VLOOKUP($A2,'Diccionario SFC'!$A$9:$C$12,3,0)`
- `CA2`: `=W2`
- `CB2`: `=X2`
- `CC2`: `=Y2`
- `CD2`: `=INDEX($CX:$CX,MATCH(VLOOKUP($CE2,$DB:$DC,2,0),$CY:$CY,0))`
- `CE2`: `=IF(AL2="SOF",AL2&"1",AL2)`
- `CF2`: `=IF(AND(CE2="SOF1",AN2=0),0,IF(AND(CE2="SOF1",AN2<>0),_xlfn.CONCAT(ROUND(AN2,2),"%"),IF(AND(CE2="IBR",AN2=0),0,IF(AND(CE2="IBR",AN2<>0),_xlfn.CONCAT(ROUND(AN2,2),"%"),IFERROR(AN2/100,0)))))`
- `CG2`: `=+VLOOKUP($AO2,$DD:$DE,2,0)`
- `CH2`: `=+IF(CG2="Trimestral","TV",IF(CG2="semestral","SV",IF(CG2="anual","AV",IF(CG2="mensual","MV","Vnto"))))`
- `CI2`: `=+AB2`
- `CJ2`: `=IF($AA2="COU","UVR",$AA2)`
- `CK2`: `=BB2`
- `CL2`: `=INDEX($CX:$CX,MATCH(VLOOKUP($CM2,$DB:$DC,2,0),$CY:$CY,0))`
- `CM2`: `=+IF(AP2="SOF",AP2&"1",AP2)`
- `CN2`: `=+IF(AND(CM2="SOF1",AR2=0),0,IF(AND(CM2="SOF1",AR2<>0),_xlfn.CONCAT(ROUND(AR2,2),"%"),IF(AND(CM2="IBR",AR2=0),0,IF(AND(CM2="IBR",AR2<>0),_xlfn.CONCAT(ROUND(AR2,2),"%"),IFERROR(AR2/100,0)))))`
- `CO2`: `=+VLOOKUP($AS2,$DD:$DE,2,0)`
- `CP2`: `=+IF(CO2="Trimestral","TV",IF(CO2="semestral","SV",IF(CO2="anual","AV",IF(CO2="mensual","MV","Vnto"))))`
- `CQ2`: `=+AD2`
- `CR2`: `=+IF($AC2="COU","UVR",$AC2)`
- `CS2`: `=BC2`
- `CT2`: `=+CK2-CS2`
- `CU2`: `=+VLOOKUP($F2,$DK$1:$DL$20,2,0)`

## Dic BD:BE (R2 TDPE/FINDI map)  [Diccionario SFC] valores
- r2:  | Especiales NEMO | Debe estar
- r3: S | ISIN | NEMO
- r4: P | COR42PT00014 | FONINMOBAL
- r5: A | COV34PT00120 | PEI
- r6: M | CORJ8PA00013 | HCOLSEL
- r7: T | CORB6PA00015 | ICOLCAP
- r8:  | CL0002892397 | NUAMCO
- r9:  | CA1348083025 | CNEC
- r10:  | COT80PT00036 | COT80PT00036
- r11:  | CORB5PT05263 | FONDARRER

## Dic BG:BH (AS2 valor nominal flag)  [Diccionario SFC] valores
- r1:  | Especiales VM X Nominal
- r2:  |  | EXCEPCIÓN
- r3:  | ISIN | CONTROL
- r4:  | MX0MGO0000H9 | 1
- r5:  | MX0MGO0000P2 | 1

## Dic BJ:BK (X2 moneda)  [Diccionario SFC] valores
- r2:  | Especiales Moneda | EXCEPCIÓN
- r3:  | ISIN | CONTROL
- r4:  | IE00B53QG562 | EUR

## Dic BB:BC (AO2)  [Diccionario SFC] valores
- r3:  | SEMESTRAL | S
- r4:  | AL VENCIMIENTO | P
- r5:  | ANUAL | A
- r6:  | MENSUAL | M
- r7:  | TRIMESTRAL | T

## VECTOR M:P (monedas/FX)  [VECTOR] valores
- r3: Moneda | ICO | Valor COP
- r4: COP | COP | 1
- r5: USD | USD | =VLOOKUP($Q5,$U:$Y,5,FAL
- r6: EUR | EUR | =VLOOKUP($Q6,$U:$Y,5,FAL
- r7: CAD | CAD | =VLOOKUP($Q7,$U:$Y,5,FAL
- r8: MXN | MXN | =VLOOKUP($Q8,$U:$Y,5,FAL
- r9: COU | UVR | =IF(Parámetros!$I$17=TRU
- r10: GBP | GBP | =VLOOKUP($Q10,$U:$Y,5,FA
- r11: CLP | CLP | =VLOOKUP($Q11,$U:$Y,5,FA
- r12: JPY | JPY | =VLOOKUP($Q12,$U:$Y,5,FA
- r13: BRL | BRL | =VLOOKUP($Q13,$U:$Y,5,FA
- r14: PEN | PEN | =VLOOKUP(Parámetros!$C$5
- r15: AUD | AUD | =VLOOKUP($Q15,$U:$Y,5,FA
