import * as $ from 'jquery';
import * as moment from 'moment';
import { IHttpClientResponse } from '../models/IHttpClientResponse';
import { IValidationField } from '../models/IValidationField';
import { IValidationResult } from '../models/IValidationResult';
import { Date } from './types';

/** utility function to check null or undefined */
export const isNullOrUndefined = (value: any) => value === null || value === undefined;

/** utility function to check null or empty */
export const isNullOrEmpty = (value: any) => value === null || value === "";

/** utility function to check null or empty or undefined*/
export const isNullOrUndefinedOrEmpty = (value: any) => isNullOrUndefined(value) || isNullOrEmpty(value);

/** utility function to check obj having property */
export const getProperty = (obj: any, prop: string) => (!isNullOrEmpty(obj) && !isNullOrUndefined(obj)) ? obj.hasOwnProperty(prop) ? obj[prop] : "" : "";

export function toLocaleLongDateString(date: Date) {
  return moment(date).format('LL');
}

export function toLocaleShortDateString(date: Date) {
  return moment(date).format('ll');
}

export function pad(value: string, length: number, valueToAppend: string) {
  if (value.length < length) {
    var append = "";
    for (var index = 0; index < length; index++) {
      append += valueToAppend
    }
    var s = append + value;
    var final = s.substr(s.length - length);
    return final;
  } else {
    return value;
  }
}

export function isObjectsEqual(o1: any, o2: any): boolean {
  return (!isNullOrUndefined(o1) && !isNullOrUndefined(o1)) && typeof o1 === 'object' && Object.keys(o1).length > 0
    ? Object.keys(o1).length === Object.keys(o2).length
    && Object.keys(o1).every(p => isObjectsEqual(o1[p], o2[p]))
    : o1 === o2;
}

export function isArrayObjectsEqual(a1: any[], a2: any[]) {
  return a1.length === a2.length && a1.every((o, idx) => isObjectsEqual(o, a2[idx]));
}

export function getDate(format?: string, action?: Date.Action,unit?: Date.Unit, amount?: number) {
  let m = moment();
  (action == "subtract") ? m.subtract(amount, unit) : m.add(amount, unit);
  return m.format(format);
}

export function findIndex(items: any[], key: string, value: string) {
  var index = -1;
  items.some(function (el, i) {
    if (el[key] == value) {
      index = i;
      return true;
    }
  });
  return index;
}

export function getObjectFromArray(arr: any[], key: string, value: string) {
  var object: any = {};
  arr.some(obj => {
    if (obj[key] == value) {
      object = { ...obj };
      return true;
    }
  });
  return object;
}

export function isTwoArraySame(arr1: string[] | number[], arr2: string[] | number[]): boolean {
  if (arr1.length == arr2.length) {
    for (var i = 0; i < arr1.length; i++) {
      if (arr1[i] == arr2[i]) {
        // do nothing
      } else {
        return false;
      }
    }
    // here you checked all array, you know that each element is
    // same, because if it wouldn't, it would return false already
    // so you can return true now
    return true;
  } else {
    return false;
  }
}

export function error(res: IHttpClientResponse, webpartName: string, showAlert: boolean = false) {
  if (!res.ok) {
    console.log(res);
    if (showAlert)
      alert(webpartName + ": " + res.error.message.value);
  }
}
/**
 * Adds a value to a date
 *
 * @param date The date to which we will add units, done in local time
 * @param interval The name of the interval to add, one of: ['year', 'quarter', 'month', 'week', 'day', 'hour', 'minute', 'second']
 * @param units The amount to add to date of the given interval
 *
 * http://stackoverflow.com/questions/1197928/how-to-add-30-minutes-to-a-javascript-date-object
 */
export function dateAdd(date, interval, units) {
  var ret = new Date(date.toString()); // don't change original date
  switch (interval.toLowerCase()) {
    case "year":
      ret.setFullYear(ret.getFullYear() + units);
      break;
    case "quarter":
      ret.setMonth(ret.getMonth() + 3 * units);
      break;
    case "month":
      ret.setMonth(ret.getMonth() + units);
      break;
    case "week":
      ret.setDate(ret.getDate() + 7 * units);
      break;
    case "day":
      ret.setDate(ret.getDate() + units);
      break;
    case "hour":
      ret.setTime(ret.getTime() + units * 3600000);
      break;
    case "minute":
      ret.setTime(ret.getTime() + units * 60000);
      break;
    case "second":
      ret.setTime(ret.getTime() + units * 1000);
      break;
    default:
      ret = undefined;
      break;
  }
  return ret;
}

/**
 * Get the local file as an array buffer.
 *@param fileElementId String that specifies the element ID.
*/
export function getFileBuffer(fileElementId: string) {
  var fileInput: any = $('#' + fileElementId);
  var deferred = $.Deferred();
  var reader = new FileReader();
  reader.onloadend = function (e: any) {
    deferred.resolve(e.target.result);
  }
  reader.onerror = function (e: any) {
    deferred.reject(e.target.error);
  }
  reader.readAsArrayBuffer(fileInput[0].files[0]);
  return deferred.promise();
}
/**
 * @param {string} latitude
 * @param {string} longitude
 * @returns
 * @memberof spservices
 */
export async function getGeoLocationName(latitude: number, longitude: number) {
  try {
    const apiUrl = `https://nominatim.openstreetmap.org/reverse?format=json&lat=${latitude}&lon=${longitude}&zoom=18&addressdetails=1`;
    const results = await $.ajax({
      url: apiUrl,
      type: 'GET',
      dataType: 'json',
      headers: {
        'content-type': 'application/json;charset=utf-8',
        'accept': 'application/json;odata=nometadata',
      }
    });

    if (results) {
      return results;
    }
  } catch (error) {
    return Promise.reject(error);
  }
}

/**
 * Get the generated column from sp fields result for datatable.
 *@param fieldResult Sharepoint fields result.
*/
export function generateDataTableColumn(fieldResult: any[]) {
  let columns = [];
  columns = fieldResult.map(field => {
    let obj = {
      title: field.Title,
      data: field.InternalName
    };
    return obj;
  });

  return columns;
}

export function validateFields(fields: IValidationField[], highlightInvalidBorder: boolean = false) {

  var validFields: IValidationField[] = [];
  var inValidFields: IValidationField[] = [];
  var result: IValidationResult = undefined;
  fields.forEach((field) => {
    document.getElementById(field.Id).classList.remove("border-danger");
    if (field.Type == "text") {
      const textElem = (<HTMLInputElement>document.getElementById(field.Id));
      isNullOrEmpty(textElem.value) || isNullOrUndefined(textElem.value) ? inValidFields.push(field) : validFields.push(field);
    } else if (field.Type == "file") {
      const fileElem = (<HTMLInputElement>document.getElementById(field.Id));
      fileElem.files.length == 0 ? inValidFields.push(field) : validFields.push(field);
    } else if (field.Type == "url") {
      const urlElem = (<HTMLInputElement>document.getElementById(field.Id));
      !validURL(urlElem.value) ? inValidFields.push(field) : validFields.push(field);
    } else if (field.Type == "select") {
      const selElem = (<HTMLInputElement>document.getElementById(field.Id));
      (selElem.value == "0" || selElem.value == "-1") ? inValidFields.push(field) : validFields.push(field);
    }
  });

  result = {
    IsValid: inValidFields.length > 0 ? false : true,
    ValidFields: validFields,
    InValidFields: inValidFields
  }

  if (highlightInvalidBorder) {
    inValidFields.forEach(field => {
      document.getElementById(field.Id).classList.add("border-danger");
    })
  }


  return result;
}

export function validURL(url: string) {
  var expression = new RegExp("^(http|https|ftp)\://([a-zA-Z0-9\.\-]+(\:[a-zA-Z0-9\.&amp;%\$\-]+)*@)*((25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9])|([a-zA-Z0-9\-]+\.)*[a-zA-Z0-9\-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(\:[0-9]+)*(/($|[a-zA-Z0-9\.\,\?\'\\\+&amp;%\$#\=~_\-]+))*$");
  var urlregex = new RegExp("^(http|https|ftp)\://([a-zA-Z0-9\.\-]+(\:[a-zA-Z0-9\.&amp;%\$\-]+)*@)*((25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9])|([a-zA-Z0-9\-]+\.)*[a-zA-Z0-9\-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(\:[0-9]+)*(/($|[a-zA-Z0-9\.\,\?\'\\\+&amp;%\$#\=~_\-]+))*$");
  //var res = url.match(/(http(s)?:\/\/.)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)/g);
  return expression.test(url) && urlregex.test(url) ? true : false;
}

export function initiateSlickPagination(sliderElemID: string, sliderContent: string) {

  // Selectors from HTML file
  var sliderID: any = $(sliderElemID);
  sliderContent = sliderContent;

  // slick carousel plugin initialize with options
  sliderID.find(sliderContent).not('.slick-initialized').slick();

  // call navigate to next slide function
  this.navigate_to_next_slide(sliderID, sliderContent);
  // call navigate to prev slide function
  this.navigate_to_prev_slide(sliderID, sliderContent);
  // call get total no of slides function
  this.get_total_no_of_slides(sliderID, sliderContent);
  // call after slider change function
  this.after_slider_change(sliderID, sliderContent);
}

export function navigate_to_next_slide(sliderID, sliderContent) {
  // click function
  sliderID.find('.slick-pag-next').on('click', function () {
    // go to next slide
    sliderID.find(sliderContent).slick('slickNext');
    // get the current slide
    var getCurrentSlide = sliderID.find(sliderContent).slick('slickCurrentSlide');
    // append the current slide
    sliderID.find('.slick-currentSlide').text(getCurrentSlide + 1);
  });
}

export function navigate_to_prev_slide(sliderID, sliderContent) {
  // click function
  sliderID.find('.slick-pag-prev').on('click', function () {
    // go to prev slide
    sliderID.find(sliderContent).slick('slickPrev');
    // get the current slide
    var getCurrentSlide = sliderID.find(sliderContent).slick('slickCurrentSlide');
    // append the current slide
    sliderID.find('.slick-currentSlide').text(getCurrentSlide + 1);
  });
}

export function get_total_no_of_slides(sliderID, sliderContent) {
  // get slick object
  var getSlideCount = sliderID.find(sliderContent).slick('getSlick');
  // append total number of slides
  sliderID.find('.slick-slideCount').text(getSlideCount.slideCount);
}

export function after_slider_change(sliderID, sliderContent) {
  // On after slide change
  sliderID.find(sliderContent).on('afterChange', function (event, slick, currentSlide, nextSlide) {
    // append the current slide
    sliderID.find('.slick-currentSlide').text(currentSlide + 1);
  });
}

var prevBtnInnerHtml: string = "";
export function btnLoader(btn: string, bindLoader: boolean, otherBtnsToDisable: string[] = []) {
  var btnElem = (document.querySelector(btn) as HTMLButtonElement);
  const btnInnerHtml = '<span class="spinner-border spinner-border-sm" style="float: left; margin: 4px 5px 0 0;" role="status" aria-hidden="true"></span>Loading...';
  if (bindLoader) {
    prevBtnInnerHtml = btnElem.innerHTML;
    btnElem.innerHTML = btnInnerHtml;
    btnElem.disabled = true;
    if (otherBtnsToDisable.length > 0)
      document.querySelectorAll(otherBtnsToDisable.join(",")).forEach((btn: HTMLButtonElement) => btn.disabled = true);
  }
  else {
    btnElem.innerHTML = prevBtnInnerHtml;
    btnElem.disabled = false;
    if (otherBtnsToDisable.length > 0)
      document.querySelectorAll(otherBtnsToDisable.join(",")).forEach((btn: HTMLButtonElement) => btn.disabled = false);
  }
}

export function getSPIcon(extension: string, title: string = "", isFolder: boolean = false) {
  let imgHtml = "";
  let icon = "";
  extension = isFolder ? "folder" : extension;
  switch (extension) {
    case "folder":
      imgHtml = '<img border="0" alt="folder" title="' + title + '" src="/_layouts/15/images/folder.gif">';
      icon = "/_layouts/15/images/folder.gif";
      break;
    case "png":
      imgHtml = '<img border="0" alt="folder" title="' + title + '" src="/_layouts/15/images/icpng.gif">';
      icon = "/_layouts/15/images/icpng.gif";
      break;
    case "jpg":
    case "jpeg":
      imgHtml = '<img border="0" alt="folder" title="' + title + '" src="/_layouts/15/images/icjpg.gif">';
      icon = "/_layouts/15/images/icjpg.gif";
      break;
    case "txt":
      imgHtml = '<img border="0" alt="txt" title="' + title + '" src="/_layouts/15/images/ictxt.gif">';
      icon = "/_layouts/15/images/ictxt.gif";
      break;
    case "pdf":
      imgHtml = '<img border="0" alt="Excel workbook" title="' + title + '" src="/_layouts/15/images/icpdf.png">';
      icon = "/_layouts/15/images/icpdf.png";
      break;
    case "xls":
      imgHtml = '<img border="0" alt="Excel workbook" title="' + title + '" src="/_layouts/15/images/icxls.png">';
      icon = "/_layouts/15/images/icxls.png";
      break;
    case "xlsx":
      imgHtml = '<img border="0" alt="Excel workbook" title="' + title + '" src="/_layouts/15/images/icxlsx.png">';
      icon = "/_layouts/15/images/icxlsx.png";
      break;
    case "doc":
      imgHtml = '<img border="0" alt="doc file" title="' + title + '" src="/_layouts/15/images/icdoc.png">';
      icon = "/_layouts/15/images/icdoc.png";
      break;
    case "docx":
      imgHtml = '<img border="0" alt="doc file" title="' + title + '" src="/_layouts/15/images/icdocx.png">';
      icon = "/_layouts/15/images/icdocx.png";
      break;
    case "ppt":
      imgHtml = '<img border="0" alt="ppt file" title="' + title + '" src="/_layouts/15/images/icppt.png">';
      icon = "/_layouts/15/images/icppt.png";
      break;
    case "pptx":
      imgHtml = '<img border="0" alt="ppt file" title="' + title + '" src="/_layouts/15/images/icpptx.png">';
      icon = "/_layouts/15/images/icpptx.png";
      break;
    default:
      imgHtml = '<img border="0" alt="default" title="' + title + '" src="/_layouts/15/images/icgen.gif">';
      icon = "/_layouts/15/images/icgen.gif";
  }

  return { html: imgHtml, icon };
}

export function downloadBlob(blob: Blob, filename: string) {
  // Create an object URL for the blob object
  const url = URL.createObjectURL(blob);

  // Create a new anchor element
  const a = document.createElement('a');

  // Set the href and download attributes for the anchor element
  // You can optionally set other attributes like `title`, etc
  // Especially, if the anchor element will be attached to the DOM
  a.href = url;
  a.download = filename || 'download';

  // Click handler that releases the object URL after the element has been clicked
  // This is required for one-off downloads of the blob content
  const clickHandler = () => {
    setTimeout(() => {
      URL.revokeObjectURL(url);
      removeEventListener('click', clickHandler);
    }, 150);
  };

  // Add the click event listener on the anchor element
  // Comment out this line if you don't want a one-off download of the blob content
  a.addEventListener('click', clickHandler, false);

  // Programmatically trigger a click on the anchor element
  // Useful if you want the download to happen automatically
  // Without attaching the anchor element to the DOM
  // Comment out this line if you don't want an automatic download of the blob content
  a.click();

  // Return the anchor element
  // Useful if you want a reference to the element
  // in order to attach it to the DOM or use it in some other way
  return a;
}

export function validateWebpartProps(props){
  let isValidProps = true;
  Object.keys(props).forEach(propName => {
    if(props[propName] == undefined)
        isValidProps = false;
  });
  if(!isValidProps)
    alert("Web part properties are not properly configured to load form data. Edit the form and configure.");
  return isValidProps;
}