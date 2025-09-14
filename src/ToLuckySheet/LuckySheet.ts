﻿import { IluckyImageBorder,IluckyImageCrop,IluckyImageDefault,IluckyImages,IluckySheetCelldata,IluckySheetCelldataValue,IMapluckySheetborderInfoCellForImp,IluckySheetborderInfoCellValue,IluckySheetborderInfoCellValueStyle,IFormulaSI,IluckySheetRowAndColumnLen,IluckySheetRowAndColumnHidden,IluckySheetSelection,IcellOtherInfo,IformulaList,IformulaListItem, IluckysheetHyperlink, IluckysheetHyperlinkType, IluckysheetDataVerification} from "./ILuck";
import { debug } from '../utils/debug';
import {LuckySheetCelldata} from "./LuckyCell";
import { IattributeList } from "../ICommon";
import {getXmlAttibute, getColumnWidthPixel, fromulaRef,getRowHeightPixel,getcellrange,generateRandomIndex,getPxByEMUs, getMultiSequenceToNum, getTransR1C1ToSequence, getPeelOffX14, getMultiFormulaValue} from "../common/method";
import {borderTypes, COMMON_TYPE2, DATA_VERIFICATION_MAP, DATA_VERIFICATION_TYPE2_MAP, worksheetFilePath} from "../common/constant";
import { ReadXml, IStyleCollections, Element,getColor } from "./ReadXml";
import { LuckyFileBase,LuckySheetBase,LuckyConfig,LuckySheetborderInfoCellForImp,LuckySheetborderInfoCellValue,LuckysheetCalcChain,LuckySheetConfigMerge } from "./LuckyBase";
import {ImageList} from "./LuckyImage";
import dayjs from "dayjs";
import { LuckyCondition } from "./LuckyCondition";
import { LuckyVerification } from "./LuckyVerification";
import { LuckFilter } from "./luckyFilter";
import { LuckyFreezen } from './luckyFreezen'
import { ChartImageGroup } from './LuckyChart'

export class LuckySheet extends LuckySheetBase {

    private readXml:ReadXml
    private sheetFile:string
    private isInitialCell:boolean
    private styles:IStyleCollections
    private sharedStrings:Element[]
    private mergeCells:Element[]
    private calcChainEles:Element[]
    private sheetList:IattributeList
    private cellImages: Element[]

    private imageList:ImageList

    private formulaRefList:IFormulaSI

    constructor(sheetName:string, sheetId:string, sheetOrder:number,isInitialCell:boolean=false, allFileOption:any){
        //Private
        super();
        this.isInitialCell = isInitialCell;
        this.readXml = allFileOption.readXml;
        this.sheetFile = allFileOption.sheetFile;
        this.styles = allFileOption.styles;
        this.sharedStrings = allFileOption.sharedStrings;
        this.calcChainEles = allFileOption.calcChain;
        this.sheetList = allFileOption.sheetList;
        this.imageList = allFileOption.imageList;
        this.hide = allFileOption.hide;
        this.cellImages = allFileOption.cellImages;

        //Output
        this.name = sheetName;
        this.index = sheetId;
        this.order = sheetOrder.toString();
        this.config = new LuckyConfig();
        this.celldata = [];
        
        // Handle empty sheets (when sheetFile is null)
        if(!this.sheetFile) {
            // Set defaults for empty sheet
            this.showGridLines = "1";
            this.status = "0";
            this.zoomRatio = 1;
            this.defaultColWidth = 73;
            this.defaultRowHeight = 19;
            this.row = 84;
            this.column = 60;
            return;
        }
        
        this.mergeCells = this.readXml.getElementsByTagName("mergeCells/mergeCell", this.sheetFile);
        let clrScheme = this.styles["clrScheme"] as Element[];
        let sheetView = this.readXml.getElementsByTagName("sheetViews/sheetView", this.sheetFile);

        let showGridLines = "1", tabSelected="0", zoomScale = "100", activeCell = "A1";
        if(sheetView.length>0){
            let attrList = sheetView[0].attributeList;
            showGridLines = getXmlAttibute(attrList, "showGridLines", "1");
            tabSelected = getXmlAttibute(attrList, "tabSelected", "0");
            zoomScale = getXmlAttibute(attrList, "zoomScale", "100");
            // let colorId = getXmlAttibute(attrList, "colorId", "0");
            let selections = sheetView[0].getInnerElements("selection");
            if(selections!=null && selections.length>0){
                activeCell = getXmlAttibute(selections[0].attributeList, "activeCell", "A1");
                let range:IluckySheetSelection = getcellrange(activeCell);
                this.luckysheet_select_save = [];
                this.luckysheet_select_save.push(range);
            }

            let pane = sheetView[0].getInnerElements("pane");
            if (pane?.length > 0) {
                this.freezen = new LuckyFreezen(pane[0])
            }
        }
        this.showGridLines = showGridLines;
        this.status = tabSelected;
        this.zoomRatio = parseInt(zoomScale)/100;

        let tabColors = this.readXml.getElementsByTagName("sheetPr/tabColor", this.sheetFile);
        if(tabColors!=null && tabColors.length>0){
            let tabColor = tabColors[0], attrList = tabColor.attributeList;
            // if(attrList.rgb!=null){
                let tc = getColor(tabColor, this.styles, "b");
                this.color = tc;
            // }
        }

        let sheetFormatPr = this.readXml.getElementsByTagName("sheetFormatPr", this.sheetFile);
        let defaultColWidth, defaultRowHeight;
        if(sheetFormatPr.length>0){
            let attrList = sheetFormatPr[0].attributeList;
            defaultColWidth = getXmlAttibute(attrList, "defaultColWidth", "9.21");
            defaultRowHeight = getXmlAttibute(attrList, "defaultRowHeight", "19");
        }

        this.defaultColWidth = getColumnWidthPixel(parseFloat(defaultColWidth));
        this.defaultRowHeight = getRowHeightPixel(parseFloat(defaultRowHeight));


        this.generateConfigColumnLenAndHidden();
        let cellOtherInfo:IcellOtherInfo =  this.generateConfigRowLenAndHiddenAddCell();
        
        if(this.calcChain==null){
            this.calcChain = [];
        }

        let formulaListExist:IformulaList={};
        for(let c=0;c<this.calcChainEles.length;c++){
            let calcChainEle = this.calcChainEles[c], attrList = calcChainEle.attributeList;
            if(attrList.i!=sheetId){
                continue;
            }

            let r = attrList.r , i = attrList.i, l = attrList.l, s = attrList.s, a = attrList.a, t = attrList.t;

            let range = getcellrange(r);
            let chain = new LuckysheetCalcChain();
            chain.r = range.row[0];
            chain.c = range.column[0];
            chain.index = this.index;
            this.calcChain.push(chain);
            formulaListExist["r"+r+"c"+c] = null;
        }
        

        // Process shared formulas
        if(this.formulaRefList!=null){
            for(let key in this.formulaRefList){
                let funclist = this.formulaRefList[key];
                let mainFunc = funclist["mainRef"], mainCellValue = mainFunc.cellValue;
                let formulaTxt = mainFunc.fv;
                let mainR = mainCellValue.r, mainC = mainCellValue.c;
                // let refRange = getcellrange(ref);
                for(let name in funclist){
                    if(name == "mainRef"){
                        continue;
                    }

                    let funcValue = funclist[name], cellValue = funcValue.cellValue;
                    if(cellValue==null){
                        continue;
                    }
                    let r = cellValue.r, c = cellValue.c;

                    let func = formulaTxt;
                    let offsetRow = r - mainR, offsetCol = c - mainC;

                    if(offsetRow > 0){
                        func = "=" + fromulaRef.functionCopy(func, "down", offsetRow);
                    }
                    else if(offsetRow < 0){
                        func = "=" + fromulaRef.functionCopy(func, "up", Math.abs(offsetRow));
                    }

                    if(offsetCol > 0){
                        func = "=" + fromulaRef.functionCopy(func, "right", offsetCol);
                    }
                    else if(offsetCol < 0){
                        func = "=" + fromulaRef.functionCopy(func, "left", Math.abs(offsetCol));
                    }

                    (cellValue.v as IluckySheetCelldataValue ).f = func;
                    
                    //添加共享公式链
                    let chain = new LuckysheetCalcChain();
                    chain.r = cellValue.r;
                    chain.c = cellValue.c;
                    chain.index = this.index;
                    this.calcChain.push(chain);
                }
            }
        }


        //There may be formulas that do not appear in calcChain
        for(let key in cellOtherInfo.formulaList){
            if(!(key in formulaListExist)){
                let formulaListItem = cellOtherInfo.formulaList[key];
                let chain = new LuckysheetCalcChain();
                chain.r = formulaListItem.r;
                chain.c = formulaListItem.c;
                chain.index = this.index;
                this.calcChain.push(chain);
            }
        }

        const conditionList = this.readXml.getElementsByTagName("conditionalFormatting", this.sheetFile)
        const  extLstCondition =
        this.readXml.getElementsByTagName(
          "extLst/ext/x14:conditionalFormattings/x14:conditionalFormatting",
          this.sheetFile
        ) || [];
        
        const extLstRule = extLstCondition?.map(condition => {
            const sqref = this.readXml.getElementsByTagName("xm:sqref", condition.value, false)?.[0]
            return this.readXml.getElementsByTagName("x14:cfRule", condition.value, false).map(d => ({
                ...d,
                parentAttribute: { sqref: sqref?.value },
                isExtLst: true,
                extLst: undefined as any
            }))
        })?.flat() || [];

        if (conditionList?.length) {
            const ruleList = conditionList.map(condition => {
                return this.readXml.getElementsByTagName("cfRule", condition.value, false)?.map(d => ({
                    ...d,
                    parentAttribute: condition.attributeList,
                    extLst: extLstRule.find((d: any) => d.parentAttribute.sqref === condition.attributeList?.sqref)
                }))
            })?.flat().filter(Boolean).concat(extLstRule?.filter((d: any) => conditionList.findIndex(condition => condition.attributeList.sqref === d.parentAttribute.sqref) === -1)) || [];
            
            this.conditionalFormatting = ruleList.map((d: any ) => new LuckyCondition(d, this.readXml, this.styles));
            // debug.log(ruleList, allFileOption, this.conditionalFormatting)
        }
        // debug.log(allFileOption)
        const filter = new LuckFilter(this.readXml, this.sheetFile)
        if (filter.ref) this.filter = filter;
      
        // dataVerification config
        this.dataVerification = this.generateConfigDataValidations();
        this.dataVerificationList = this.generateConfigDataValidationsList();
        // debug.log('dataVerificationList ---->', this.dataVerificationList)

        // hyperlink config
        this.hyperlink = this.generateConfigHyperlinks();
      
        // sheet hide
        this.hide = this.hide;

        if(this.mergeCells!=null){
            for(let i=0;i<this.mergeCells.length;i++){
                let merge = this.mergeCells[i], attrList = merge.attributeList;
                let ref = attrList.ref;
                if(ref==null){
                    continue;
                }
                let range = getcellrange(ref);
                let mergeValue = new LuckySheetConfigMerge();
                mergeValue.r = range.row[0];
                mergeValue.c = range.column[0];
                mergeValue.rs = range.row[1]-range.row[0]+1;
                mergeValue.cs = range.column[1]-range.column[0]+1;
                if(this.config.merge==null){
                    this.config.merge = {};
                }
                this.config.merge[range.row[0] + "_" + range.column[0]] = mergeValue;
            }
        }

        let drawingFile = allFileOption.drawingFile, drawingRelsFile = allFileOption.drawingRelsFile;
        if(drawingFile!=null && drawingRelsFile!=null){
            this.getImageBaseInfo(drawingFile, drawingRelsFile)
        } 
    }

    private getImageBaseInfo = (drawingFile: string, drawingRelsFile: string): any => {
        let twoCellAnchors = this.readXml.getElementsByTagName("xdr:twoCellAnchor", drawingFile);
        let oneCellAnchors = this.readXml.getElementsByTagName("xdr:oneCellAnchor", drawingFile);
        twoCellAnchors=[...twoCellAnchors,...oneCellAnchors];
        if(twoCellAnchors!=null && twoCellAnchors.length>0){
            for(let i=0;i<twoCellAnchors.length;i++){
                let twoCellAnchor = twoCellAnchors[i];
                
                let xdrFroms = twoCellAnchor.getInnerElements("xdr:from"), xdrTos = twoCellAnchor.getInnerElements("xdr:to");

                if(xdrFroms!=null && xdrFroms.length>0){
                    let xdrFrom = xdrFroms[0], xdrTo, xdrExt;
                    if(xdrTos){
                        xdrTo = xdrTos[0];
                    }else{
                        xdrExt = twoCellAnchor.getInnerElements("xdr:ext")[0]
                    }
                    let imageObject: any = {};
                    
                    let xdr_graphicFrame = twoCellAnchor.getInnerElements("xdr:graphicFrame");
                    if (xdr_graphicFrame) {
                        imageObject = this.getGraphic(twoCellAnchor, drawingRelsFile)
                    }
                    let xdr_pic = twoCellAnchor.getInnerElements("xdr:pic");
                    if (xdr_pic) {
                        imageObject = this.getImage(twoCellAnchor, drawingRelsFile)
                    }

                    // let imageObject = xdr_graphicFrame ? this.getGraphic(twoCellAnchor, drawingRelsFile) : this.getImage(twoCellAnchor, drawingRelsFile)

                    let x_n =0,y_n = 0;
                    let cx_n = 0, cy_n = 0;

                    imageObject.fromCol = this.getXdrValue(xdrFrom.getInnerElements("xdr:col"));
                    imageObject.fromColOff = getPxByEMUs(this.getXdrValue(xdrFrom.getInnerElements("xdr:colOff")));
                    imageObject.fromRow= this.getXdrValue(xdrFrom.getInnerElements("xdr:row"));
                    imageObject.fromRowOff = getPxByEMUs(this.getXdrValue(xdrFrom.getInnerElements("xdr:rowOff")));
                    if(xdrTo){
                        imageObject.toCol = this.getXdrValue(xdrTo.getInnerElements("xdr:col"));
                        imageObject.toColOff = getPxByEMUs(this.getXdrValue(xdrTo.getInnerElements("xdr:colOff")));
                        imageObject.toRow = this.getXdrValue(xdrTo.getInnerElements("xdr:row"));
                        imageObject.toRowOff = getPxByEMUs(this.getXdrValue(xdrTo.getInnerElements("xdr:rowOff")));
                    }else{
                        let a = xdrExt.attributeList
                        cx_n = getPxByEMUs(parseInt(a.cx)),cy_n = getPxByEMUs(parseInt(a.cy));
                        imageObject.toCol = imageObject.fromCol;
                        imageObject.toColOff = Number(imageObject.fromColOff)+cx_n;
                        imageObject.toRow = imageObject.fromRow;
                        imageObject.toRowOff = Number(imageObject.fromRowOff)+cy_n;
                    }
                    imageObject.originWidth = cx_n;
                    imageObject.originHeight = cy_n;
                    
                    imageObject.isFixedPos = false;
                    imageObject.fixedLeft = 0;
                    imageObject.fixedTop = 0;

                    let imageBorder:IluckyImageBorder = {
                        color: "#000",
                        radius: 0,
                        style: "solid",
                        width: 0
                    }
                    imageObject.border = imageBorder;

                    let imageCrop:IluckyImageCrop = {
                        height: cy_n,
                        offsetLeft: 0,
                        offsetTop: 0,
                        width: cx_n
                    }
                    imageObject.crop = imageCrop;

                    let imageDefault:IluckyImageDefault = {
                        height: cy_n,
                        left: x_n,
                        top: y_n,
                        width: cx_n
                    }
                    imageObject.default = imageDefault;
                    if(this.images==null){
                        this.images = {};
                    }

                    if (imageObject.id) {
                        this.images[imageObject.id || generateRandomIndex("image")] = imageObject;
                    }
                }
            }
        }
        return null
    }
    private getImage = (twoCellAnchor: Element, drawingRelsFile: string) => {
        let xdr_blipfills = twoCellAnchor.getInnerElements("a:blip");
        let editAs = getXmlAttibute(twoCellAnchor.attributeList, "editAs", "twoCell");
        if (xdr_blipfills!=null && xdr_blipfills.length>0) {
            var xdr_blipfill = xdr_blipfills[0];
            let rembed = getXmlAttibute(xdr_blipfill.attributeList, "r:embed", null);

            let imageObject: any = this.getBase64ByRid(rembed, drawingRelsFile);

            if(editAs=="absolute"){
                imageObject.type = "3";
            }
            else if(editAs=="oneCell"){
                imageObject.type = "2";
            }
            else{
                imageObject.type = "1";
            }
            return imageObject
        }
        return {}
    }
    private getGraphic = (twoCellAnchor: Element, drawingRelsFile: string) => {
        try {
            const xdr_graphicFrames = twoCellAnchor.getInnerElements("xdr:graphicFrame");
            if (xdr_graphicFrames.length) {
                const xdr_graphicFrame = xdr_graphicFrames[0];
                const chartImageGroup = new ChartImageGroup({
                    graphicFrame: xdr_graphicFrame,
                    readXml: this.readXml,
                    drawingRelsFile,
                    styles: this.styles,
                })
                const imageObject = chartImageGroup.image;
                if (chartImageGroup.chart) {
                    if(this.charts==null){
                        this.charts = [];
                    }
                    this.charts.push(chartImageGroup.chart)
                }
                return imageObject;
            }
        } catch (error) {
            debug.warn('Failed to process chart/graphic, skipping:', error);
            // Return empty object to continue processing
        }
        return {};
    }

    private getXdrValue(ele:Element[]):number{
        if(ele==null || ele.length==0){
            return null;
        }

        return parseInt(ele[0].value);
    }

    private getBase64ByRid(rid:string, drawingRelsFile:string){
        let Relationships = this.readXml.getElementsByTagName("Relationships/Relationship", drawingRelsFile);

        if(Relationships!=null && Relationships.length>0){
            for(let i=0;i<Relationships.length;i++){
                let Relationship = Relationships[i];
                let attrList = Relationship.attributeList;
                let Id = getXmlAttibute(attrList, "Id", null);
                let src = getXmlAttibute(attrList, "Target", null);
                if(Id == rid){
                    src = src.replace(/\.\.\//g, "");
                    src = "xl/" + src;
                    let imgage = this.imageList.getImageByName(src);
                    return imgage;
                }
            }
        }

        return {};
    }

    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
    private generateConfigColumnLenAndHidden(){
        let cols = this.readXml.getElementsByTagName("cols/col", this.sheetFile);
        for(let i=0;i<cols.length;i++){
            let col = cols[i], attrList = col.attributeList;
            let min = getXmlAttibute(attrList, "min", null);
            let max = getXmlAttibute(attrList, "max", null);
            let width = getXmlAttibute(attrList, "width", null);
            let hidden = getXmlAttibute(attrList, "hidden", null);
            let customWidth = getXmlAttibute(attrList, "customWidth", null);


            if(min==null || max==null){
                continue;
            }

            let minNum = parseInt(min)-1, maxNum=parseInt(max)-1, widthNum=parseFloat(width);
            
            for(let m=minNum;m<=maxNum;m++){
                if(width!=null){
                    if(this.config.columnlen==null){
                        this.config.columnlen = {};
                    }
                    this.config.columnlen[m] = getColumnWidthPixel(widthNum);
                }

                if(hidden=="1"){
                    if(this.config.colhidden==null){
                        this.config.colhidden = {};
                    }
                    this.config.colhidden[m] = 0;
                    
                }

                if(customWidth!=null){
                    if(this.config.customWidth==null){
                        this.config.customWidth = {};
                    }
                    this.config.customWidth[m] = 1;
                }
            } 
        }
    }

    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
    private generateConfigRowLenAndHiddenAddCell():IcellOtherInfo{
        // Remove verbose logging
        let rows = this.readXml.getElementsByTagName("sheetData/row", this.sheetFile);
        let cellOtherInfo:IcellOtherInfo = {};
        let formulaList:IformulaList = {};
        cellOtherInfo.formulaList = formulaList;
        for(let i=0;i<rows.length;i++){
            let row = rows[i], attrList = row.attributeList;
            let rowNo = getXmlAttibute(attrList, "r", null);
            let height = getXmlAttibute(attrList, "ht", null);
            let hidden = getXmlAttibute(attrList, "hidden", null);
            let customHeight = getXmlAttibute(attrList, "customHeight", null);

            if(rowNo==null){
                continue;
            }

            let rowNoNum = parseInt(rowNo) - 1;
            if(height!=null){
                let heightNum = parseFloat(height);
                if(this.config.rowlen==null){
                    this.config.rowlen = {};
                }
                this.config.rowlen[rowNoNum] = getRowHeightPixel(heightNum);
            }

            if(hidden=="1"){
                if(this.config.rowhidden==null){
                    this.config.rowhidden = {};
                }
                this.config.rowhidden[rowNoNum] = 0;
                
            }

            if(customHeight!=null){
                if(this.config.customHeight==null){
                    this.config.customHeight = {};
                }
                this.config.customHeight[rowNoNum] = 1;
            }


            if(this.isInitialCell){
                let cells = row.getInnerElements("c");
                for(let key in cells){
                    let cell = cells[key];
                    const cellSize = this.getCellSize(cell);
                    let cellValue = new LuckySheetCelldata(
                        cell, 
                        cellSize,
                        this.styles, 
                        this.sharedStrings, 
                        this.mergeCells,
                        this.sheetFile, 
                        this.cellImages, 
                        this.imageList, 
                        this.readXml
                    );
                    if(cellValue._borderObject!=null){
                        if(this.config.borderInfo==null){
                            this.config.borderInfo = [];
                        }
                        this.config.borderInfo.push(cellValue._borderObject);
                        delete cellValue._borderObject;
                    }
                    
                    // let borderId = cellValue._borderId;
                    // if(borderId!=null){
                    //     let borders = this.styles["borders"] as Element[];
                    //     if(this.config._borderInfo==null){
                    //         this.config._borderInfo = {};
                    //     }
                    //     if( borderId in this.config._borderInfo){
                    //         this.config._borderInfo[borderId].cells.push(cellValue.r + "_" + cellValue.c);
                    //     }
                    //     else{
                    //         let border = borders[borderId];
                    //         let borderObject = new LuckySheetborderInfoCellForImp();
                    //         borderObject.rangeType = "cellGroup";
                    //         borderObject.cells = [];
                    //         let borderCellValue = new LuckySheetborderInfoCellValue();
                            
                    //         let lefts = border.getInnerElements("left");
                    //         let rights = border.getInnerElements("right");
                    //         let tops = border.getInnerElements("top");
                    //         let bottoms = border.getInnerElements("bottom");
                    //         let diagonals = border.getInnerElements("diagonal");

                    //         let left = this.getBorderInfo(lefts);
                    //         let right = this.getBorderInfo(rights);
                    //         let top = this.getBorderInfo(tops);
                    //         let bottom = this.getBorderInfo(bottoms);
                    //         let diagonal = this.getBorderInfo(diagonals);

                    //         let isAdd = false;
                    //         if(left!=null && left.color!=null){
                    //             borderCellValue.l = left;
                    //             isAdd = true;
                    //         }

                    //         if(right!=null && right.color!=null){
                    //             borderCellValue.r = right;
                    //             isAdd = true;
                    //         }

                    //         if(top!=null && top.color!=null){
                    //             borderCellValue.t = top;
                    //             isAdd = true;
                    //         }

                    //         if(bottom!=null && bottom.color!=null){
                    //             borderCellValue.b = bottom;
                    //             isAdd = true;
                    //         }

                    //         if(isAdd){
                    //             borderObject.value = borderCellValue;
                    //             this.config._borderInfo[borderId] = borderObject;
                    //         }

                    //     }
                    // }
                    if(cellValue._formulaType=="shared"){
                        // Remove verbose logging
                        if(this.formulaRefList==null){
                            this.formulaRefList = {};
                        }

                        if(this.formulaRefList[cellValue._formulaSi]==null){
                            this.formulaRefList[cellValue._formulaSi] = {}
                        }
                        
                        const currentCellRef = String.fromCharCode(65 + cellValue.c) + (cellValue.r + 1);
                        const formula = cellValue.v ? (cellValue.v as IluckySheetCelldataValue).f : 'unknown';
                        debug.log(`🔧 [SharedFormula] Collecting cell ${currentCellRef} with SI=${cellValue._formulaSi}, formula="${formula}", hasRef=${!!cellValue._fomulaRef}`);

                        let fv;
                        if(cellValue.v!=null){
                            fv = (cellValue.v as IluckySheetCelldataValue).f;
                            // Fix =+ prefix in shared formulas before expansion
                            if(fv && fv.startsWith('=+')) {
                                fv = '=' + fv.substring(2);
                            }
                        }

                        let refValue = {
                            t:cellValue._formulaType,
                            ref:cellValue._fomulaRef,
                            si:cellValue._formulaSi,
                            fv:fv,
                            cellValue:cellValue
                        }

                        if(cellValue._fomulaRef!=null){
                            this.formulaRefList[cellValue._formulaSi]["mainRef"] = refValue;
                        }
                        else{
                            this.formulaRefList[cellValue._formulaSi][cellValue.r+"_"+cellValue.c] = refValue;
                        }

                        // debug.log(refValue, this.formulaRefList);
                    }

                    //There may be formulas that do not appear in calcChain
                    if(cellValue.v!=null && (cellValue.v as IluckySheetCelldataValue).f!=null){
                        let formulaCell:IformulaListItem = {
                            r:cellValue.r,
                            c:cellValue.c
                        }
                        cellOtherInfo.formulaList["r"+cellValue.r+"c"+cellValue.c] = formulaCell;
                    }

                    this.celldata.push(cellValue);
                }
                
            }
        }

        return cellOtherInfo;
    }
    private generateConfigDataValidationsList() {
        let rows = this.readXml.getElementsByTagName(
            "dataValidations/dataValidation",
            this.sheetFile
          );
          let extLst =
            this.readXml.getElementsByTagName(
              "extLst/ext/x14:dataValidations/x14:dataValidation",
              this.sheetFile
            ) || [];
          
          rows = rows.concat(extLst);
          return rows.map(d => new LuckyVerification(d, extLst)).filter(d => d.uid)
    }
    /**
     * luckysheet config of dataValidations
     * 
     * @returns {IluckysheetDataVerification} - dataValidations config
     */
    private generateConfigDataValidations(): IluckysheetDataVerification {
      
      let rows = this.readXml.getElementsByTagName(
        "dataValidations/dataValidation",
        this.sheetFile
      );
      let extLst =
        this.readXml.getElementsByTagName(
          "extLst/ext/x14:dataValidations/x14:dataValidation",
          this.sheetFile
        ) || [];
      
      rows = rows.concat(extLst);
  
      let dataVerification: IluckysheetDataVerification = {};
  
      for (let i = 0; i < rows.length; i++) {
        let row = rows[i];
        let attrList = row.attributeList;
        let formulaValue = row.value;
  
        let type = getXmlAttibute(attrList, "type", null);
        if(!type) {
            continue;
        }
        let operator = "",
            sqref = "",
            sqrefIndexArr: string[] = [],
            valueArr: string[] = [];
        let _prohibitInput =
          getXmlAttibute(attrList, "allowBlank", null) !== "1" ? false : true;
        
        // x14 processing
        const formulaReg = new RegExp(/<x14:formula1>|<xm:sqref>/g)
        if (formulaReg.test(formulaValue) && extLst?.length >= 0) {
          operator = getXmlAttibute(attrList, "operator", null);
          const peelOffData = getPeelOffX14(formulaValue);
          sqref = peelOffData?.sqref;
          sqrefIndexArr = getMultiSequenceToNum(sqref);
          valueArr = getMultiFormulaValue(peelOffData?.formula);
        } else {
          operator = getXmlAttibute(attrList, "operator", null);
          sqref = getXmlAttibute(attrList, "sqref", null);
          sqrefIndexArr = getMultiSequenceToNum(sqref);
          valueArr = getMultiFormulaValue(formulaValue);
        }

        let _type = DATA_VERIFICATION_MAP[type];
        let _type2 = null;
        let _value1: string | number = valueArr?.length >= 1 ? valueArr[0] : "";
        let _value2: string | number = valueArr?.length === 2 ? valueArr[1] : "";
        let _hint = getXmlAttibute(attrList, "prompt", null);
        let _hintShow = _hint ? true : false
  
        const matchType = COMMON_TYPE2.includes(_type) || !DATA_VERIFICATION_TYPE2_MAP[_type] ? "common" : _type;
        _type2 = operator
          ? DATA_VERIFICATION_TYPE2_MAP[matchType][operator]
          : "bw";
        
        // mobile phone number processing
        if (
          _type === "text_content" &&
          (_value1?.includes("LEN") || _value1?.includes("len")) &&
          _value1?.includes("=11")
        ) {
          _type = "validity";
          _type2 = "phone";
        }

        // date processing
        if (_type === "date") {
          const D1900 = new Date(1899, 11, 30, 0, 0, 0);
          _value1 = dayjs(D1900)
            .clone()
            .add(Number(_value1), "day")
            .format("YYYY-MM-DD");
          _value2 = dayjs(D1900)
            .clone()
            .add(Number(_value2), "day")
            .format("YYYY-MM-DD");
        }
        
        // checkbox and dropdown processing
        if (_type === "checkbox" || _type === "dropdown") {
          _type2 = null;
        }
        
        // dynamically add dataVerifications
        for (const ref of sqrefIndexArr) {
          dataVerification[ref] = {
            type: _type,
            type2: _type2,
            value1: _value1,
            value2: _value2,
            checked: false,
            remote: false,
            prohibitInput: _prohibitInput,
            hintShow: _hintShow,
            hintText: _hint
          };
        }
      }
  
      return dataVerification;
    }
  
    /**
     * luckysheet config of hyperlink
     * 
     * @returns {IluckysheetHyperlink} - hyperlink config
     */
    private generateConfigHyperlinks(): IluckysheetHyperlink {
      let rows = this.readXml.getElementsByTagName(
        "hyperlinks/hyperlink",
        this.sheetFile
      );
      let hyperlink: IluckysheetHyperlink = {};
      for (let i = 0; i < rows.length; i++) {
        let row = rows[i];
        let attrList = row.attributeList;
        let ref = getXmlAttibute(attrList, "ref", null),
            refArr = getMultiSequenceToNum(ref),
            _display = getXmlAttibute(attrList, "display", null),
            _address = getXmlAttibute(attrList, "location", null),
            _tooltip = getXmlAttibute(attrList, "tooltip", null);
        let _type: IluckysheetHyperlinkType = _address ? "internal" : "external";
  
        // external hyperlink
        if (!_address) {
          let rid = attrList["r:id"];
          let sheetFile = this.sheetFile;
          let relationshipList = this.readXml.getElementsByTagName(
            "Relationships/Relationship",
            `xl/worksheets/_rels/${sheetFile.replace(worksheetFilePath, "")}.rels`
          );
  
          const findRid = relationshipList?.find(
            (e) => e.attributeList["Id"] === rid
          );

          if (findRid) {
            _address = findRid.attributeList["Target"];
            _type = findRid.attributeList[
              "TargetMode"
            ]?.toLocaleLowerCase() as IluckysheetHyperlinkType;
          }
        }

        // match R1C1 - use a more efficient pattern
        const addressReg = /^[^!]*!R([\d$])+C([\d$])*$/
        if (addressReg.test(_address)) {
          _address = getTransR1C1ToSequence(_address);
        }
        
        // dynamically add hyperlinks
        for (const ref of refArr) {
          hyperlink[ref] = {
            linkAddress: _address,
            linkTooltip: _tooltip || "",
            linkType: _type,
            display: _display || "",
          };
        }
      }
      
      return hyperlink;
    }

    // private getBorderInfo(borders:Element[]):LuckySheetborderInfoCellValueStyle{
    //     if(borders==null){
    //         return null;
    //     }

    //     let border = borders[0], attrList = border.attributeList;
    //     let clrScheme = this.styles["clrScheme"] as Element[];
    //     let style:string = attrList.style;
    //     if(style==null || style=="none"){
    //         return null;
    //     }

    //     let colors = border.getInnerElements("color");
    //     let colorRet = "#000000";
    //     if(colors!=null){
    //         let color = colors[0];
    //         colorRet = getColor(color, clrScheme);
    //     }

    //     let ret = new LuckySheetborderInfoCellValueStyle();
    //     ret.style = borderTypes[style];
    //     ret.color = colorRet;

    //     return ret;
    // }
    private getCellSize = (cell: Element) => {
        let attrList = cell.attributeList;
        let r = attrList.r, s = attrList.s, t = attrList.t;
        let range = getcellrange(r);

        const row = range.row[0];
        const col = range.column[0];

        const width = this.config.columnlen && this.config.columnlen[col] ? this.config.columnlen[col] : this.defaultColWidth;
        const height = this.config.rowlen && this.config.rowlen[row]? this.config.rowlen[row] : this.defaultRowHeight;
        return {
            width,
            height
        }
    }
}
