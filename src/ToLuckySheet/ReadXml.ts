import {IuploadfileList, IattributeList, stringToNum} from "../ICommon";
import {indexedColors}  from "../common/constant";
import {LightenDarkenColor}  from "../common/method";
import { debug } from "../utils/debug";


class xmloperation {
    // Escape problematic characters in XML attributes to prevent regex issues
    private escapeXmlAttributes(xmlString: string): string {
        // DISABLED: This regex was causing heap exhaustion on large files (>700KB)
        // The regex /([\w:]+)="([^"]*)"/ with capturing groups causes catastrophic backtracking
        // Properly formatted Excel XML shouldn't have unescaped > in attribute values anyway
        // If needed in future, use a streaming parser or chunked approach instead
        return xmlString;
    }
    
    // Restore escaped characters back to original
    private unescapeXmlAttributes(xmlString: string): string {
        // DISABLED: Since we're not escaping anymore, just return the string as-is
        return xmlString;
    }
    
    /**
    * @param tag Search xml tag name , div,title etc.
    * @param file Xml string
    * @return Xml element string 
    */
    protected getElementsByOneTag(tag:string, file:string):string[]{
        // Escape problematic characters in the XML before parsing
        const escapedFile = this.escapeXmlAttributes(file);
        
        //<a:[^/>: ]+?>.*?</a:[^/>: ]+?>
        let readTagReg;

        // For very large files, avoid catastrophic regex backtracking by processing in chunks
        // Lower threshold to catch more problematic files
        const isLargeFile = escapedFile.length > 100000;

        if(isLargeFile) {
            console.log(`[XML] Large file detected (${escapedFile.length} chars) for tag "${tag}", processing in chunks`);

            // For very large files, use a different approach to avoid regex catastrophic backtracking
            // Instead of regex, use a simpler string search approach
            const allMatches: string[] = [];
            const tags = tag.indexOf("|") > -1 ? tag.split("|") : [tag];

            for(const t of tags) {
                let searchPos = 0;
                const openTag = `<${t}`;
                const closeTag = `</${t}>`;
                const selfClose = `/>`;

                // For extremely large files (>10MB), limit the number of matches to prevent memory issues
                const maxMatches = escapedFile.length > 10000000 ? 1000 : -1;
                let matchCount = 0;

                while(searchPos < escapedFile.length) {
                    // Check if we've hit the match limit for extremely large files
                    if(maxMatches > 0 && matchCount >= maxMatches) {
                        console.log(`[XML] Hit match limit (${maxMatches}) for tag "${t}" in large file`);
                        break;
                    }

                    // Find the next occurrence of this tag
                    const tagStart = escapedFile.indexOf(openTag, searchPos);
                    if(tagStart === -1) break;

                    // Check if it's a self-closing tag first
                    const nextGt = escapedFile.indexOf(">", tagStart);
                    if(nextGt !== -1) {
                        const tagContent = escapedFile.substring(tagStart, nextGt + 1);

                        // Check if it's self-closing
                        if(tagContent.endsWith("/>")) {
                            allMatches.push(tagContent);
                            searchPos = nextGt + 1;
                            matchCount++;
                            continue;
                        }

                        // Look for the closing tag
                        const closePos = escapedFile.indexOf(closeTag, nextGt);
                        if(closePos !== -1) {
                            // For extremely large individual tags (>1MB), skip them to avoid memory issues
                            const tagLength = closePos + closeTag.length - tagStart;
                            if(tagLength > 1000000) {
                                console.log(`[XML] Skipping extremely large tag (${tagLength} chars) for "${t}"`);
                                searchPos = closePos + closeTag.length;
                                continue;
                            }

                            const fullTag = escapedFile.substring(tagStart, closePos + closeTag.length);
                            allMatches.push(fullTag);
                            searchPos = closePos + closeTag.length;
                            matchCount++;
                        } else {
                            // No closing tag found, might be malformed
                            searchPos = nextGt + 1;
                        }
                    } else {
                        break;
                    }
                }
            }

            let ret = allMatches.length > 0 ? allMatches : null;
            if(ret==null){
                return [];
            }
            else{
                // Unescape the matched results before returning
                return ret.map(match => this.unescapeXmlAttributes(match));
            }
        }

        // Original approach for smaller files
        if(tag.indexOf("|")>-1){
            let tags = tag.split("|"), tagsRegTxt="";
            for(let i=0;i<tags.length;i++){
                let t = tags[i];
                tagsRegTxt += "|<"+ t +" [^>]+?[^/]>[\\s\\S]*?</"+ t +">|<"+ t +" [^>]+?/>|<"+ t +">[\\s\\S]*?</"+ t +">|<"+ t +"/>";
            }
            tagsRegTxt = tagsRegTxt.substr(1, tagsRegTxt.length);
            readTagReg = new RegExp(tagsRegTxt, "g");
        }
        else{
            readTagReg = new RegExp("<"+ tag +" [^>]+?[^/]>[\\s\\S]*?</"+ tag +">|<"+ tag +" [^>]+?/>|<"+ tag +">[\\s\\S]*?</"+ tag +">|<"+ tag +"/>", "g");
        }

        let ret = escapedFile.match(readTagReg);
        if(ret==null){
            return [];
        }
        else{
            // Unescape the matched results before returning
            return ret.map(match => this.unescapeXmlAttributes(match));
        }
    }

    protected getElementByTagLink(tag:string, file:string): string[]{
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(file, "text/xml");

        let tagVal: globalThis.Element[] = Array.from(xmlDoc.children);
        if(tag.indexOf("/")>-1){
            let tags = tag.split("/");
            for (let index = 0; index < tags.length; index++) {
                const element = tags[index];
                const i = tagVal.findIndex(d => Array.from(d.children).findIndex(d => d.tagName === element) > -1);
                if (i === -1 && index <= tags.length - 1) {
                    return [];
                }
                tagVal = Array.from(tagVal[i].children).filter(d => d.tagName === element)
            }
        } else {
            tagVal = Array.from(tagVal[0].children).filter(d => d.tagName === tag)
        }
        const serializer = new XMLSerializer();
        return tagVal.map(d => serializer.serializeToString(d));
    }
}

export class ReadXml extends xmloperation{
    originFile:IuploadfileList
    constructor(files:IuploadfileList){
        super();
        this.originFile = files;
    }
    /**
    * @param path Search xml tag group , div,title etc.
    * @param fileName One of uploadfileList, uploadfileList is file group, {key:value}
    * @return Xml element calss
    */
    getElementsByTagName(path:string, fileName:string, isFile: boolean = true): Element[]{
        
        let file = this.getFileByName(fileName);
        if (!isFile) file = fileName;
        let pathArr = path.split("/"), ret:string[] | string;
        for(let key in pathArr){
            let path = pathArr[key];
            if(ret==undefined){
                ret = this.getElementsByOneTag(path,file);
            }
            else{
                if(ret instanceof Array){
                    let items:string[]=[];
                    for(let key in ret){
                        let item = ret[key];
                        items = items.concat(this.getElementsByOneTag(path,item));
                    }
                    ret = items;
                }
                else{
                    ret = this.getElementsByOneTag(path,ret);
                }
            }
        }

        let elements:Element[] = [];

        for(let i=0;i<ret.length;i++){
            let ele = new Element(ret[i]);
            elements.push(ele);
        }

        return elements;
    }

    getElementsByTagNameLink(path:string, fileName:string, isFile: boolean = true): Element[]{
        let file = this.getFileByName(fileName);
        if (!isFile) file = fileName;
        const ret = this.getElementByTagLink(path, file);
        let elements:Element[] = [];

        for(let i=0;i<ret.length;i++){
            let ele = new Element(ret[i]);
            elements.push(ele);
        }

        return elements;
    }

    /**
    * @param name One of uploadfileList's name, search for file by this parameter
    * @retrun Select a file from uploadfileList
    */
    private getFileByName(name:string):string{
        for(let fileKey in this.originFile){
            if(fileKey.indexOf(name)>-1){
                return this.originFile[fileKey];
            }
        }
        return "";
    }

    
}

export class Element extends xmloperation {
    elementString:string
    attributeList:IattributeList
    value:string
    container:string
    constructor(str:string){
        super();
        this.elementString = str;
        this.setValue();
        // For very large container strings, use a safer approach
        let attrList: string[] | null;
        if (this.container.length > 50000) {
            // Use a safer approach for large strings
            attrList = [];
            let pos = 0;
            while (pos < this.container.length) {
                // Find next attribute pattern
                const eqPos = this.container.indexOf('="', pos);
                if (eqPos === -1) break;

                // Find attribute name start
                let nameStart = eqPos - 1;
                while (nameStart >= 0 && /[a-zA-Z0-9_:]/.test(this.container[nameStart])) {
                    nameStart--;
                }
                nameStart++;

                // Find attribute value end
                const valueEnd = this.container.indexOf('"', eqPos + 2);
                if (valueEnd === -1) break;

                const attr = this.container.substring(nameStart, valueEnd + 1);
                attrList.push(attr);
                pos = valueEnd + 1;
            }
            if (attrList.length === 0) attrList = null;
        } else {
            const readAttrReg = new RegExp('[a-zA-Z0-9_:]*?=".*?"', "g");
            attrList = this.container.match(readAttrReg);
        }
        this.attributeList = {};
        if(attrList!=null){
            for(let key in attrList){
                let attrFull = attrList[key];
                // let al= attrFull.split("=");
                if(attrFull.length==0){
                    continue;
                }
                let attrKey = attrFull.substr(0, attrFull.indexOf('='));
                let attrValue = attrFull.substr(attrFull.indexOf('=') + 1);
                if(attrKey==null || attrValue==null ||attrKey.length==0 || attrValue.length==0){
                    continue;
                }
                // Extract the attribute value (remove surrounding quotes)
                let unescapedValue = attrValue.substr(1, attrValue.length-2);
                // No need to unescape since we're not escaping anymore
                this.attributeList[attrKey] = unescapedValue;
            }
        }
    }

    /**
    * @param name Get attribute by key in element
    * @return Single attribute
    */
    get(name:string):string|number|boolean{
        return this.attributeList[name];
    }

    /**
    * @param tag Get elements by tag in elementString
    * @return Element group
    */
    getInnerElements(tag:string):Element[]{
        let ret = this.getElementsByOneTag(tag,this.elementString);
        let elements:Element[] = [];

        for(let i=0;i<ret.length;i++){
            let ele = new Element(ret[i]);
            elements.push(ele);
        }

        if(elements.length==0){
            return null;
        }
        return elements;
    }

    getInnerElementsTagLink(tag: string): Element[]{
        const ret = this.getElementByTagLink(tag, this.elementString);
        let elements:Element[] = [];

        for(let i=0;i<ret.length;i++){
            let ele = new Element(ret[i]);
            elements.push(ele);
        }
        if(elements.length==0){
            return null;
        }

        return elements;
    }

    /**
    * @desc get xml dom value and container, <container>value</container>
    */
    private setValue(){
        let str = this.elementString;
        if(str.substr(str.length-2, 2)=="/>"){
            this.value = "";
            this.container = str;
        }
        else{
            let firstTag = this.getFirstTag();

            // For very large strings, process more carefully to avoid heap exhaustion
            if (str.length > 50000) {
                console.log(`[XML setValue] Large string detected (${str.length} chars), using safe parsing`);
                // Instead of changing logic, just limit the regex scope
                // Find the closing tag position first to limit regex range
                const closeTag = "</" + firstTag + ">";
                const closeTagPos = str.indexOf(closeTag);

                if (closeTagPos !== -1) {
                    // Only apply regex to the portion up to and including the close tag
                    const limitedStr = str.substring(0, closeTagPos + closeTag.length);
                    const firstTagReg = new RegExp("(<"+ firstTag +" [^>]+?[^/]>)([\\s\\S]*?)</"+ firstTag +">|(<"+ firstTag +">)([\\s\\S]*?)</"+ firstTag +">", "g");
                    let result = firstTagReg.exec(limitedStr);
                    if (result != null) {
                        if(result[1]!=null){
                            this.container = result[1];
                            this.value = result[2];
                        }
                        else{
                            this.container = result[3];
                            this.value = result[4];
                        }
                    }
                } else {
                    // No closing tag found, treat as self-closing
                    this.container = str;
                    this.value = "";
                }
            } else {
                // Original regex approach for smaller strings
                const firstTagReg = new RegExp("(<"+ firstTag +" [^>]+?[^/]>)([\\s\\S]*?)</"+ firstTag +">|(<"+ firstTag +">)([\\s\\S]*?)</"+ firstTag +">", "g");
                let result = firstTagReg.exec(str);
                if (result != null) {
                    if(result[1]!=null){
                        this.container = result[1];
                        this.value = result[2];
                    }
                    else{
                        this.container = result[3];
                        this.value = result[4];
                    }
                }
            }
        }
    }

    /**
    * @desc get xml dom first tag, <a><b></b></a>, get a
    */
    private getFirstTag(){
        let str = this.elementString;
        let firstTag = str.substr(0, str.indexOf(' '));
        if(firstTag=="" || firstTag.indexOf(">")>-1){
            firstTag = str.substr(0, str.indexOf('>'));
        }
        firstTag = firstTag.substr(1,firstTag.length);
        return firstTag;
    }
}


export interface IStyleCollections {
    [index:string]:Element[] | IattributeList
}

function combineIndexedColor(indexedColorsInner:Element[], indexedColors:IattributeList):IattributeList{
    let ret:IattributeList = {};
    if(indexedColorsInner==null || indexedColorsInner.length==0){
        return indexedColors;
    }
    for(let key in indexedColors){
        let value = indexedColors[key], kn = parseInt(key);
        let inner = indexedColorsInner[kn];
        if(inner==null){
            ret[key] = value;
        }
        else{
            let rgb = inner.attributeList.rgb;
            ret[key] = rgb;
        }
    }

    return ret;
}

//clrScheme:Element[]
export function getColor(color:Element, styles:IStyleCollections , type:string="g"){
    let attrList = color.attributeList;
    let clrScheme = styles["clrScheme"] as Element[];
    let indexedColorsInner = styles["indexedColors"] as Element[];
    let mruColorsInner = styles["mruColors"];
    let indexedColorsList = combineIndexedColor(indexedColorsInner, indexedColors);
    let indexed = attrList.indexed, rgb = attrList.rgb, theme = attrList.theme, tint = attrList.tint;
    let bg;
    if(indexed!=null){
        let indexedNum = parseInt(indexed);
        bg = indexedColorsList[indexedNum];
        if(bg!=null){
            bg = bg.substring(bg.length-6, bg.length);
            bg = "#"+bg;
        }
    }
    else if(rgb!=null){
        rgb = rgb.substring(rgb.length-6, rgb.length);
        bg = "#"+rgb;
    }
    else if(theme!=null){
        let themeNum = parseInt(theme);
        if(themeNum==0){
            themeNum = 1;
        }
        else if(themeNum==1){
            themeNum = 0;
        }
        else if(themeNum==2){
            themeNum = 3;
        }
        else if(themeNum==3){
            themeNum = 2;
        }
        let clrSchemeElement = clrScheme[themeNum];
        if(clrSchemeElement!=null){
            let clrs = clrSchemeElement.getInnerElements("a:sysClr|a:srgbClr");
            if(clrs!=null){
                let clr = clrs[0];
                let clrAttrList = clr.attributeList;
                // debug.log(clr.container, );
                if(clr.container.indexOf("sysClr")>-1){
                    // if(type=="g" && clrAttrList.val=="windowText"){
                    //     bg = null;
                    // }
                    // else if((type=="t" || type=="b") && clrAttrList.val=="window"){
                    //     bg = null;
                    // }                    
                    // else 
                    if(clrAttrList.lastClr!=null){
                        bg = "#" + clrAttrList.lastClr;
                    }
                    else if(clrAttrList.val!=null){
                        bg = "#" + clrAttrList.val;
                    }

                }
                else if(clr.container.indexOf("srgbClr")>-1){
                    // debug.log(clrAttrList.val);
                    bg = "#" + clrAttrList.val;
                }
            }
        }
        
    }
    
    if(tint!=null){
        let tintNum = parseFloat(tint);
        if(bg!=null){
            bg = LightenDarkenColor(bg, tintNum);
        }
    }

    return bg;
}


/** 
 * @dom xml attribute object
 * @attr attribute name
 * @d if attribute is null, return default value 
 * @return attribute value
*/
export function getlineStringAttr(frpr:Element, attr:string):string{
    let attrEle = frpr.getInnerElements(attr), value;

    if(attrEle!=null && attrEle.length>0){
        if(attr=="b" || attr=="i" || attr=="strike"){
            value = "1";
        }
        else if(attr=="u"){
            let v = attrEle[0].attributeList.val;
            if(v=="double"){
                value =  "2";
            }
            else if(v=="singleAccounting"){
                value =  "3";
            }
            else if(v=="doubleAccounting"){
                value =  "4";
            }
            else{
                value = "1";
            }
        }
        else if(attr=="vertAlign"){
            let v = attrEle[0].attributeList.val;
            if(v=="subscript"){
                value = "1";
            }
            else if(v=="superscript"){
                value = "2";
            }
        }
        else{
            value = attrEle[0].attributeList.val;
        }
        
    }

    return value;
}