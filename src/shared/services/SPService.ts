/*! *****************************************************************************
Author : 
EMail  : 
***************************************************************************** */

import {
  SPHttpClient,
  SPHttpClientResponse,
  IHttpClientOptions,
  IDigestCache,
  DigestCache,
} from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as utilities from '../utils/utilities';
import * as moment from 'moment';
import * as constants from "../utils/constants";
import * as enums from '../utils/enums';
import * as $ from "jquery";
import { IEventData } from "../models/IEventData";
import { IHttpClientResponse } from "../models/IHttpClientResponse";
import { IHttpClientError } from "../models/IHttpClientError";
import { SPComponentLoader } from '@microsoft/sp-loader';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import { IUser } from "../models/IUser";
import { IUrl } from "../models/IUrl";
import { ILegacyPageContext } from "../models/ILegacyPageContext";
import { IEmailTemplate } from "../models/IEmailTemplate";
import { IFormNoConfiguration } from "../models/IFormNoConfiguration";
import { MSGraphClient } from '@microsoft/sp-http';
require("../styles/custom.scss");
//require('SPServices');
//require('../js/jquery.SPServices.js');

if (window.location.href.indexOf("/_layouts/15/workbench.aspx")) {
  require('../../shared/styles/workbench.scss');
}

export default class SPService {
  private context: WebPartContext;
  public digest: string;
  public moment: typeof moment;
  public utils: typeof utilities;
  public constant: typeof constants;
  public enum: typeof enums;
  public rootSite: string;
  public absoluteUrl: string;
  public serverRelativeUrl: string;
  public rootURL: string;
  public pageContextInfo: ILegacyPageContext;
  /**
   * Service constructor
   * @inheritdoc
   * import SPService from '../../../utilities/SPService';
   * @example
   * const ApiHelper: SPService = new SPService(this.props.context);
   */
  constructor(_pageContext: WebPartContext) {
    this.context = _pageContext;
    this.moment = moment;
    this.utils = utilities;
    this.constant = constants;
    this.enum = enums;
    this.absoluteUrl = _pageContext.pageContext.web.absoluteUrl;
    this.serverRelativeUrl = _pageContext.pageContext.web.serverRelativeUrl;
    this.rootSite = this.absoluteUrl.indexOf("/sites/") > -1 ? this.absoluteUrl.split(_pageContext.pageContext.web.serverRelativeUrl)[0] : _pageContext.pageContext.site.absoluteUrl;
    let rootURL = this.serverRelativeUrl.split("/");
    this.rootURL = "/" + rootURL[1] + "/" + rootURL[2]
    this.pageContextInfo = _pageContext.pageContext.legacyPageContext;

    /** Load Css */
    // SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_FONTS);
    // SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_DATATABLE);
    SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_STYLE);
    SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_BOOTSTRAP);
    SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_JQUERYUI);
    SPComponentLoader.loadCss('https://fonts.googleapis.com/css?family=Lato');
    SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_FONTAWESOME);
    SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_SUMMERNOTE);
    // SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_FONTAWESOMEALL);
    SPComponentLoader.loadCss(this.absoluteUrl + this.constant.CSS_CUSTOM);
    // SPComponentLoader.loadCss(this.rootSite + this.constant.CSS_MATERIALDESIGNICONS);
    // SPComponentLoader.loadCss(this.rootSite + this.constant.CSS_SUSS);
    // SPComponentLoader.loadCss(this.rootSite + this.constant.CSS_DATATABLE);
    // SPComponentLoader.loadCss(this.rootSite + this.constant.CSS_TREEVIEW);
    // SPComponentLoader.loadCss(this.rootSite + this.constant.CSS_FONTAWESOME5);

    /** Load JS*/
    SPComponentLoader.loadScript(this.absoluteUrl + this.constant.JS_JQUERY).then(() => {
      SPComponentLoader.loadScript(this.absoluteUrl + this.constant.JS_BOOTSTRAP);
      SPComponentLoader.loadScript(this.absoluteUrl + this.constant.JS_BOOTSTRAPBUNDLE);
      // SPComponentLoader.loadScript(this.absoluteUrl + this.constant.JS_FONTAWESOME);

      // SPComponentLoader.loadScript( this.absoluteUrl + this.constant.JS_DATATABLE);
      // SPComponentLoader.loadScript( this.absoluteUrl + this.constant.JS_JQUERYUI);
      // SPComponentLoader.loadScript( this.absoluteUrl + this.constant.JS_OFFCANVAS);
      // SPComponentLoader.loadScript( this.absoluteUrl + this.constant.JS_HOVERABLECOLLAPSE);
      // SPComponentLoader.loadScript( this.absoluteUrl + this.constant.JS_MISC);
    });

    //To get the digest value for AJAX post call
    const digestCache: IDigestCache = _pageContext.serviceScope.consume(
      DigestCache.serviceKey
    );
    digestCache
      .fetchDigest(_pageContext.pageContext.web.serverRelativeUrl)
      .then((digest) => {
        this.digest = digest;
      });
  }

  /*=====================================================
          To check the folder is Exists
    =======================================================*/
  /**
   *
   * @private
   * @returns {Promise<bpolean>}
   * @memberof spservices
   */
  public async isFolderExists(serverRelativeUrl: string, siteUrl: string = "", createIfNotExists: boolean = false): Promise<boolean> {
    siteUrl = siteUrl == "" ? this.absoluteUrl : siteUrl;
    let url = siteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + serverRelativeUrl + "')/Exists";
    const exists = await this.get(url);
    let isCreated = false;
    if (exists.ok && !exists.data.value && createIfNotExists) {
      let folderUrl = `${siteUrl}/_api/web/Folders`;
      var fData = {
        ServerRelativeUrl: serverRelativeUrl
      }
      const result = await this.post(folderUrl, fData);
      isCreated = result.ok;
    }
    return exists.ok ? exists.data.value ? true : isCreated : false;
  }

  /*=====================================================
          To check the file is Exists
    =======================================================*/
  /**
   *
   * @private
   * @returns {Promise<bpolean>}
   * @memberof spservices
   */
  public async isFileExists(serverRelativeUrl: string, siteUrl: string = ""): Promise<boolean> {
    siteUrl = siteUrl == "" ? this.absoluteUrl : siteUrl;
    let url = siteUrl + "/_api/Web/GetFileByServerRelativeUrl('" + serverRelativeUrl + "')/Exists";
    const exists = await this.get(url);
    return exists.ok ? exists.data.value : false;
  }

  /*=====================================================
          To get Document Type Icon using MapToIcon
    =======================================================*/
  /**
   *
   * @private
   * @returns {Promise<any>}
   * @memberof spservices
   */
  public async getDocumentTypeIcon(fileName: string, progId: string = '', size: number = 0) {
    let url = this.absoluteUrl + "/_api/web/maptoicon(filename='" + fileName + "', progid='" + progId + "', size=" + size + ")";
    const icon = await this.get(url);
    return icon.ok ? icon.data.MapToIcon : icon.error;
  }

  /*=====================================================
            Convert UTC time to Local
    =======================================================*/
  /**
   *
   * @private
   * @returns {Promise<string>}
   * @memberof spservices
   */
  public async getLocalTime(utcTime: string | Date): Promise<string> {
    try {
      var dateIsoString: any;
      if (typeof utcTime === "string") {
        dateIsoString = utcTime;
      }
      else {
        //dateIsoString = utcTime.toISOString();
        dateIsoString = this.utils.dateAdd(utcTime, "minute", utcTime.getTimezoneOffset() * -1).toISOString();
      }
      var uri = this.absoluteUrl + "/_api/web/RegionalSettings/TimeZone/utcToLocalTime(@date)?@date='" + dateIsoString + "'";
      const localTime = await this.get(uri);
      return localTime.data.value;
    }
    catch (error) {
      return Promise.reject(error);
    }
  }

  /*=====================================================
            Convert Local time to UTC
    =======================================================*/
  /**
   *
   * @private
   * @returns {Promise<string>}
   * @memberof spservices
   */
  public async getUtcTime(localTime: string | Date): Promise<string> {
    try {
      var dateIsoString: any;
      if (typeof localTime === "string") {
        dateIsoString = localTime;
      }
      else {
        //dateIsoString = localTime.toLocaleString();
        dateIsoString = this.utils.dateAdd(localTime, "minute", localTime.getTimezoneOffset() * -1).toISOString();
      }
      var uri = this.absoluteUrl + "/_api/web/RegionalSettings/TimeZone/localTimeToUTC(@date)?@date='" + dateIsoString + "'";
      const utcTime = await this.get(uri);
      return utcTime.data.value;
    }
    catch (error) {
      return Promise.reject(error);
    }
  }

  /*=====================================================
            Uploaded Document using  Rest API
    =======================================================*/
  /**
   * You can upload files up to 2 GB with the REST API.
   * @param serverRelativeUrl Server Relative Url of the folder or library.
   * @param elementId String that specifies the ID value.
   * @param metadata Metadata for the document (optional).
   * @example
   * var serverRelativeUrl = "/sites/rootsite/subsite/shared documents",
   * var elementId = "getFile"
   */
  public async uploadFileToFolderUsingRestApi(
    serverRelativeUrl: string,
    elementId: string,
    metadata?: object
  ): Promise<any> {

    var _metadata = JSON.parse(JSON.stringify(metadata));

    // get the file name from the file input control on the page.
    var fileInput: HTMLInputElement = <HTMLInputElement>(
      document.getElementById(elementId)
    );

    if (fileInput.files.length == 0) return "File is empty";

    var fileName = fileInput.files.item(0).name;

    // get the local file as an array buffer.
    var arrayBuffer: any = await this.utils.getFileBuffer(elementId);

    // Construct the endpoint.
    var fileCollectionEndpoint = `${this.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${serverRelativeUrl}')/files/add(overwrite=true, url='${fileName}')`;

    // Send the request and return the response.
    // This call returns the SharePoint file.
    try {
      const response = await $.ajax({
        url: fileCollectionEndpoint,
        type: "post",
        data: arrayBuffer,
        processData: false,
        headers: {
          accept: "application/json;odata=verbose",
          "X-RequestDigest": this.digest,
          "content-length": arrayBuffer.byteLength,
        },
      });

      console.log(`${fileName} successfully uploaded in ${serverRelativeUrl}`);

      if (_metadata != undefined) {
        if (response.d.hasOwnProperty("ListItemAllFields")) {
          //To check the uri property

          var fileListItemUri = response.d.ListItemAllFields.__deferred.uri;
          const listItem = await this.GETByRestAPI(fileListItemUri);
          _metadata["__metadata"] = { type: `${listItem.data.d.__metadata.type}` };
          const result = await this.UPDATEByRestAPI(
            listItem.data.d.__metadata.uri,
            _metadata
          );

          return result;
        } else {
          return response;
        }
      } else {
        return response;
      }
    } catch (err) {
      console.error(err);
      return err;
    }
  }

  /*=====================================================
            Uploaded Document using  SPHttpClient
    =======================================================*/
  /**
   * You can upload files up to 2 GB with SPHttpClient.
   * @param serverRelativeUrl Server Relative Url of the folder or library.
   * @param elementId String that specifies the ID value.
   * @param metadata Metadata for the document (optional).
   * @param fileInput File input value
   * @example
   * var serverRelativeUrl = "/sites/rootsite/subsite/shared documents",
   * var elementId = "getFile"
   */
  public async uploadFileToFolder(
    serverRelativeUrl: string,
    elementId: string,
    metadata?: object,
    file?: File | undefined,
    site: string = ""
  ): Promise<any> {
    // get the file name from the file input control on the page.
    var fileInput: HTMLInputElement = <HTMLInputElement>(
      document.getElementById(elementId)
    );

    if (file == undefined)
      if (fileInput.files.length == 0) return "File is empty";

    var fileName = file != undefined ? file.name : fileInput.files.item(0).name;

    var _file = file != undefined ? file : fileInput.files.item(0);

    // Construct the endpoint.
    site = site == "" ? this.absoluteUrl : site;
    var fileCollectionEndpoint = `${site}/_api/web/GetFolderByServerRelativeUrl('${serverRelativeUrl}')/files/add(overwrite=true,url='${fileName}')`;

    // Construct headers
    const header = {
      accept: "application/json",
      "Content-type": "application/json",
    };

    const httpClientOptions: IHttpClientOptions = {
      body: _file,
      headers: header,
    };

    // Send the request and return the response.
    // This call returns the SharePoint file.
    return this.context.spHttpClient
      .post(
        fileCollectionEndpoint,
        SPHttpClient.configurations.v1,
        httpClientOptions
      )
      .then((res) => {
        return res.json().then(
          async (response) => {
            console.log(
              `${fileName} successfully uploaded in ${serverRelativeUrl}`
            );

            if (metadata != undefined) {
              var _metadata = JSON.parse(JSON.stringify(metadata));
              if (response.hasOwnProperty("@odata.id")) {
                //To check the uri property

                var fileListItemUri =
                  response["@odata.id"] + "/ListItemAllFields";
                const listItem = await this.GETByRestAPI(
                  fileListItemUri
                );
                const result = await this.update(
                  listItem.data.d.__metadata.uri,
                  _metadata
                );

                return result;
              } else {
                return response;
              }
            } else {
              return response;
            }
          },
          (error) => {
            console.error(error);
            return error;
          }
        );
      });
  }

  /*=====================================================
            update List Item using  Rest API
    =======================================================*/
  /**
   * To update the list item with the REST API.
   * @param itemUrl URI of the item to update.
   * @param metadata Metadata for the item.
   * @example
   * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items(1)",
   */
  public async UPDATEByRestAPI(
    itemUrl: string,
    metadata: object
  ): Promise<IHttpClientResponse> {
    var url = itemUrl;
    var _metadata = JSON.parse(JSON.stringify(metadata));

    try {
      //To check and add the list type name if not exists
      if (!_metadata.hasOwnProperty("__metadata")) {
        var itemMetadata = await this.GETByRestAPI(itemUrl);
        _metadata["__metadata"] = itemMetadata.data.d.__metadata;
      }

      var body = JSON.stringify(_metadata);

      // Send the request and return the promise.
      // This call does not return response content from the server.
      const response = $.ajax({
        url: url,
        type: "post",
        data: body,
        headers: {
          accept: "application/json;odata=verbose",
          "X-RequestDigest": this.digest,
          "content-Type": "application/json;odata=verbose",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
      });

      return response.then(() => {
        let response: IHttpClientResponse = {
          ok: true,
          status: 200,
          statusText: 'success',
          error: undefined,
          type: 'update',
          url: url
        }
        return response;
      }, (err) => {
        let response: IHttpClientResponse = {
          ok: false,
          status: err.status,
          statusText: err.statusText,
          error: err.responseJSON.error,
          type: 'update',
          url: url
        }
        return response;
      });

      //return response;
    }
    catch (error) {
      return Promise.reject(error);
    }
  }

  /*=====================================================
            update List Item using  SPHTTPClient
    =======================================================*/
  /**
   * To update the list item with the SPHTTPClient.
   * @param itemUrl URI of the item to update.
   * @param metadata Metadata for the item.
   * @example
   * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items(1)",
   */
  public update(itemUrl: string, metadata: object): Promise<IHttpClientResponse> {
    var url = itemUrl;
    var _metadata = JSON.parse(JSON.stringify(metadata));

    //To remove the __metadata property if exists
    if (_metadata.hasOwnProperty("__metadata")) {
      delete _metadata.__metadata;
    }

    const header = {
      Accept: "application/json;odata=nometadata",
      "Content-type": "application/json;odata=nometadata",
      "odata-version": "",
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE",
    };

    const httpClientOptions: IHttpClientOptions = {
      headers: header,
      body: JSON.stringify(_metadata),
    };

    // Send the request and return the promise.
    // This call does not return response content from the server.
    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, httpClientOptions)
      .then(
        (response: SPHttpClientResponse) => {
          let res: IHttpClientResponse = {
            ok: response.ok,
            status: response.status,
            statusText: response.statusText,
            error: undefined,
            type: response.type,
            url: response.url
          }
          if (!response.ok) {
            return response.json().then((responseJSON: any) => {
              res.error = responseJSON["odata.error"];
              return res;
            });
          } else {
            return res;
          }
        },
        (error: any) => {
          console.error(error);
          let response: IHttpClientResponse = {
            ok: false,
            status: 500,
            statusText: 'failed',
            error: error,
            type: 'update',
            url: url
          }
          return response;
        }
      );
  }

  /*=====================================================
            update List Item by id using  Rest API
    =======================================================*/
  /**
   * To update item by id
   * @param listName List name.
   * @param itemId Item Id.
   */

  public async updateItemById(listName: string, itemId: number | string, metadata: object, site: string = ""): Promise<IHttpClientResponse> {
    var id = (typeof itemId === "string") ? parseInt(itemId) : itemId;
    var _metadata = JSON.parse(JSON.stringify(metadata));
    site = site == "" ? this.absoluteUrl : site;
    var uri = site + `/_api/web/lists/getbytitle('${listName}')/items(${id})`;
    const result = await this.update(uri, _metadata);
    return result;
  }

  /*=====================================================
            Retrieve List Item using  Rest API
    =======================================================*/
  /**
   * To get the list item with the REST API..
   * @param itemUrl URI of the item to retrieve.
   * @example
   * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items",
   */
  public async GETByRestAPI(ItemUrl: string): Promise<IHttpClientResponse> {
    // Send the request and return the response.
    const response = $.ajax({
      url: ItemUrl,
      type: "get",
      headers: { accept: "application/json;odata=verbose" },
    });

    return response.then((res) => {
      if (res.d.hasOwnProperty("__next")) {
        res.d["hasNext"] = true;
      } else {
        res.d["hasNext"] = false;
      }
      let response: IHttpClientResponse = {
        ok: true,
        status: 200,
        statusText: 'success',
        data: res,
        error: undefined,
        type: 'retrieve',
        url: ItemUrl
      }
      return response;
    }, (err) => {
      let response: IHttpClientResponse = {
        ok: false,
        status: err.status,
        statusText: err.statusText,
        error: err.responseJSON.error,
        type: 'retrieve',
        url: ItemUrl
      }
      return response;
    });
  }

  /*=====================================================
            Retrieve List Item using  SPHTTPClient
    =======================================================*/
  /**
   * To retrieve the list item with the SPHTTPClient.
   * @param itemUrl URI of the item to retrieve.
   * @example
   * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items",
   */
  public get(itemUrl: string): Promise<IHttpClientResponse> {
    // Send the request and return the promise.
    // This call does not return response content from the server.
    return this.context.spHttpClient
      .get(itemUrl, SPHttpClient.configurations.v1)
      .then(
        (response: SPHttpClientResponse) => {
          return response.json().then((responseJSON) => {
            let error: IHttpClientError = undefined;
            if (!response.ok) {
              error = {
                code: responseJSON.error.code,
                message: {
                  lang: "en-US",
                  value: responseJSON.error.message
                }
              }
            }
            let res: IHttpClientResponse = {
              ok: response.ok,
              status: response.status,
              statusText: response.statusText,
              data: responseJSON,
              error: error,
              type: response.type,
              url: response.url
            }
            return res;
          }, (error: any) => {
            let err: IHttpClientResponse = {
              ok: response.ok,
              status: response.status,
              statusText: response.statusText,
              error: error,
              type: response.type,
              url: response.url
            }
            return err;
          });
        },
        (error: any) => {
          let err: IHttpClientResponse = {
            ok: false,
            status: 500,
            statusText: 'failed',
            error: error,
            type: 'post',
            url: itemUrl
          }
          return err;
        }
      );
  }

  /*=====================================================
            Add List Item using  Rest API
    =======================================================*/
  /**
   * To add the list item with the REST API.
   * @param itemUrl URI of the item to add.
   * @param metadata Metadata for the item.
   * @example
   * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items",
   */
  public async POSTByRestAPI(
    itemUrl: string,
    metadata: object = { "__metadata": "" }
  ): Promise<any> {
    var url = itemUrl;
    var _metadata = JSON.parse(JSON.stringify(metadata));

    //To check and add the list type name if not exists
    if (!_metadata.hasOwnProperty("__metadata")) {
      var itemMetadata = await this.GETByRestAPI(itemUrl);
      _metadata["__metadata"] = itemMetadata.data.d.__metadata;
    }

    var body = JSON.stringify(_metadata);

    // Send the request and return the promise.
    // This call does not return response content from the server.
    const response = await $.ajax({
      url: url,
      type: "post",
      data: body,
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": this.digest,
      },
    });

    return response;
  }

  /*=====================================================
            Add Item using  SPHTTPClient
    =======================================================*/
  /**
   * To add the list item with the SPHTTPClient.
   * @param itemUrl URI of the item to add.
   * @param metadata Metadata for the item.
   * @example
   * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items",
   */
  public post(itemUrl: string, metadata: object = {}, httpClientOptions: IHttpClientOptions = undefined): Promise<any> {
    var url = itemUrl;
    var _metadata = JSON.parse(JSON.stringify(metadata));

    //To remove the __metadata property if exists
    if (_metadata.hasOwnProperty("__metadata")) {
      delete _metadata.__metadata;
    }

    const header = {
      Accept: "application/json;odata=nometadata",
      "Content-type": "application/json;odata=nometadata",
      "odata-version": "",
    };

    if (this.utils.isNullOrUndefined(httpClientOptions)) {
      httpClientOptions = {
        headers: header,
        body: JSON.stringify(_metadata),
      };
    }

    // Send the request and return the promise.
    // This call does not return response content from the server.
    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, httpClientOptions)
      .then(
        (response: SPHttpClientResponse) => {
          console.log(`Item successfully added`);
          return response;
        },
        (error: any) => {
          console.error(error);
          return error;
        }
      );
  }

  /*=====================================================
            delete List Item using  Rest API
    =======================================================*/
  /**
   * To delete the list item with the REST API.
   * @param itemUrl URI of the item to delete.
   * @example
   * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items(1)",
   */
  public async deleteByRestAPI(itemUrl: string): Promise<any> {
    var url = itemUrl;

    // Send the request and return the promise.
    // This call does not return response content from the server.
    const response = await $.ajax({
      url: url,
      type: "post",
      headers: {
        // Accept header: Specifies the format for response data from the server.
        Accept: "application/json;odata=verbose",
        //Content-Type header: Specifies the format of the data that the client is sending to the server
        "Content-Type": "application/json;odata=verbose",
        // IF-MATCH header: Provides a way to verify that the object being changed has not been changed since it was last retrieved.
        // "IF-MATCH":"*", will overwrite any modification in the object, since it was last retrieved.
        "IF-MATCH": "*",
        //X-HTTP-Method:  The MERGE method updates only the properties of the entity , while the PUT method replaces the existing entity with a new one that you supply in the body of the post
        "X-HTTP-Method": "delete",
        // X-RequestDigest header: When you send a post request, it must include the form digest value in X-RequestDigest header
        "X-RequestDigest": this.digest,
      },
    });

    return response;
  }

  /*=====================================================
            delete List Item using  SPHTTPClient
    =======================================================*/
  /**
   * To delete the list item with the SPHTTPClient.
   * @param itemUrl URI of the item to delete.
   * @example
   * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items(1)",
   */
  public delete(itemUrl: string): Promise<any> {
    var url = itemUrl;

    const header = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': '',
      'IF-MATCH': "*",
      'X-HTTP-Method': 'DELETE'
    };

    const httpClientOptions: IHttpClientOptions = {
      headers: header,
    };

    // Send the request and return the promise.
    // This call does not return response content from the server.
    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, httpClientOptions)
      .then(
        (response: SPHttpClientResponse) => {
          console.log(`Item successfully deleted`);
          return response;
        },
        (error: any) => {
          console.error(error);
          return error;
        }
      );
  }

  /*=====================================================
            delete List Item using  SPHTTPClient
    =======================================================*/
  /**
   * To delete the list item using Id with the SPHTTPClient.
   * @param listName List name.
   * @param itemId Item Id.
   */
  public async deleteItemById(listName: string, itemId: string | number): Promise<any> {
    var id = (typeof itemId === "string") ? parseInt(itemId) : itemId;
    var uri = this.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/items(${id})`;
    const result = await this.delete(uri);
    return result;
  }

  /*=====================================================
            Upload multiple files using SPHTTPClient
    =======================================================*/
  /**
   * You can upload multiple files with the SPHTTPClient.
   * @param serverRelativeUrl Server Relative Url of the folder or library.
   * @param elementId String that specifies the ID value.
   * @param metadata Metadata for the document (optional).
   * @example
   * var serverRelativeUrl = "/sites/rootsite/subsite/shared documents",
   * var elementId = "getFile"
   */
  public async uploadMultipleFilesToFolder(
    serverRelativeUrl: string,
    elementId: string,
    metadata?: object,
    site: string = "",
    files: any[] = []
  ): Promise<any> {
    // get values from the file input and text input page controls.
    var fileInput: HTMLInputElement = <HTMLInputElement>(
      document.getElementById(elementId)
    );

    let filesArr = fileInput == null ? files : fileInput.files;

    if (filesArr.length == 0) return "File is empty";

    var fileCount = filesArr.length;
    var count: number = 0;
    var filesResponse = Array.prototype.map.call(
      filesArr,
      async (file: File) => {
        const response = await this.uploadFileToFolder(
          serverRelativeUrl,
          elementId,
          metadata,
          file,
          site
        );
        count++;
        console.log("Total file uploaded: " + count + " of " + fileCount);
        return response;
      }
    );

    const result = await Promise.all(filesResponse);
    return result;
  }

  /*=====================================================
            Retrieve Paged List Item using  Rest API
    =======================================================*/
  /**
   * To get the list item with the REST API..
   * @param selectedList List name.
   * @param selectedFields Fields to retrieve.
   * @param filterQuery Fields to filter.
   * @param orderBy Order By.
   */
  public async getPagedListItems(
    selectedList: string,
    selectedFields: any[] = [],
    filterQuery: string = "",
    orderBy: string = "",
    expand: any[] = [],
    url: string = "",
  ) {
    try {
      var selectQuery: any[] = ["Id"];
      var expandQuery: any[] = expand;
      var listItems = [];
      for (var i = 0; i < selectedFields.length; i++) {
        switch (selectedFields[i].fieldType) {
          case this.constant.FIELD_TYPE_USER:
            selectQuery.push(
              `${selectedFields[i].key}/Title,${selectedFields[i].key}/Id,${selectedFields[i].key}/Name`
            );
            expandQuery.push(selectedFields[i].key);
            break;
          case this.constant.FIELD_TYPE_LOOKUP:
            selectQuery.push(`${selectedFields[i].key}/Title`);
            expandQuery.push(selectedFields[i].key);
            break;
          case this.constant.FIELD_TYPE_ATTACHMENT:
            selectQuery.push("Attachments,AttachmentFiles");
            expandQuery.push("AttachmentFiles");
            break;
          default:
            selectQuery.push(selectedFields[i].key);
            break;
        }
      }
      var apiUri = url == "" ? this.absoluteUrl + `/_api/web/lists/getbytitle('${selectedList}')/items` : url;
      if (selectedFields.length != 0) {
        apiUri += `?$Select=${selectQuery.join()}&$expand=${expandQuery.join()}&$filter=${filterQuery}&$orderby=${orderBy}&$top=${this.constant.LIST_PAGED_LIMIT}`;
      } else {
        apiUri += `?$expand=${expandQuery.join()}&$filter=${filterQuery}&$orderby=${orderBy}&$top=${this.constant.LIST_PAGED_LIMIT}`;
      }
      var items = await this.GETByRestAPI(apiUri);
      if (items.ok) {
        listItems = items.data.d.results;
      } else {
        throw (items);
      }
      while (items.data.d.hasNext) {
        items = await this.GETByRestAPI(items.data.d.__next);
        if (items.ok) {
          listItems = [...listItems, ...items.data.d.results];
        } else {
          throw (items);
        }
      }
      return listItems;
    } catch (err) {
      Promise.reject(err);
      return listItems;
    }
  }

  /*=====================================================
                  Retrieve item by id
    =======================================================*/
  /**
   * To Retrieve item by id
   * @param listName List name.
   * @param itemId Item Id.
   */

  public async getItemById(listName: string, itemId: number | string, select: string = "", expand: string = "", site: string = ""): Promise<any> {
    var id = (typeof itemId === "string") ? parseInt(itemId) : itemId;
    site = site == "" ? this.absoluteUrl : site;
    var uri = site + `/_api/web/lists/getbytitle('${listName}')/items(${id})?$select=${select}&$expand=${expand}`;
    const result = await this.get(uri);
    var item = result.data;
    return item;
  }

  /*=====================================================
        Retrieve recurrence calendar events using SPService
    =======================================================*/
  /**
   * To get the list item with the REST API..
   * @param calendarName List name.
   */
  public async getRecurrenceEvents(calendarName: string, rowLimit: number = 0, eventPeriod: number = this.enum.CalendarEventPeriod.Year): Promise<IEventData[]> {
    var today = this.moment().format();
    var events: IEventData[] = [];
    var periodTag = "";
    switch (eventPeriod) {
      case this.enum.CalendarEventPeriod.Today:
        periodTag = "<Today />";
        break;
      case this.enum.CalendarEventPeriod.Month:
        periodTag = "<Month />";
        break;
      default:
        periodTag = "<Year />";
    }
    //select("RecurrenceID", "MasterSeriesItemID", "Id", "ID", "ParticipantsPickerId", "EventType", "Title", "Description", "EventDate", "EndDate", "Location", "Author/SipAddress", "Author/Title", "Geolocation", "fAllDayEvent", "fRecurrence", "RecurrenceData", "RecurrenceData", "Duration", "Category", "UID")
    var sp = ($() as any).SPServices;
    const calendarPromises = ($() as any).SPServices.SPGetListItemsJson({
      listName: calendarName,
      CAMLRowLimit: rowLimit,
      CAMLViewFields: "<ViewFields>" +
        "<FieldRef Name='ID' />" +
        "<FieldRef Name='Title' />" +
        "<FieldRef Name='EventDate' />" +
        "<FieldRef Name='EndDate' />" +
        "<FieldRef Name='Location' />" +
        "<FieldRef Name='Description' />" +
        "<FieldRef Name='Category' />" +
        "<FieldRef Name='fRecurrence' />" +
        "<FieldRef Name='RecurrenceData' />" +
        "<FieldRef Name='fAllDayEvent' />" +
        "</ViewFields>",
      CAMLQuery: "<Query>" +
        "<Where>" +
        "<DateRangesOverlap>" +
        "<FieldRef Name='EventDate' />" +
        "<FieldRef Name='EndDate' />" +
        "<FieldRef Name='RecurrenceID' />" +
        "<Value Type='DateTime'>" +
        periodTag +
        "</Value>" +
        "</DateRangesOverlap>" +
        "</Where>" +
        "<OrderBy>" +
        "<FieldRef Name='EventDate' />" +
        "</OrderBy>" +
        "</Query>",
      CAMLQueryOptions: "<QueryOptions>" +
        "<CalendarDate>" + today + "</CalendarDate>" +
        "<ExpandRecurrence>TRUE</ExpandRecurrence>" +
        "<RecurrenceOrderBy>TRUE</RecurrenceOrderBy>" +
        "<ViewAttributes Scope='RecursiveAll'/>" +
        "</QueryOptions>",
      mappingOverrides: {
        "ows_fAllDayEvent": {
          "mappedName": "fAllDayEvent",
          "objectType": "Boolean"
        },
        "ows_fRecurrence": {
          "mappedName": "fRecurrence",
          "objectType": "Boolean"
        }
      }
    });
    return $.when(calendarPromises).then(function () {
      var calendarEvents = this;

      $(calendarEvents.data).each(function () {
        var event: IEventData = {
          listName: calendarName,
          Id: this.ID,
          title: this.Title,
          location: this.Location != null ? this.Location : "",
          Category: this.Category,
          EventDate: this.EventDate,
          monthName: moment(this.EventDate).format("MMMM"),
          monthShortName: moment(this.EventDate).format("MMM"),
          eventDay: moment(this.EventDate).format("DD"),
          EndDate: this.EndDate,
          fAllDayEvent: this.fAllDayEvent,
          fRecurrence: this.fRecurrence,
          RecurrenceData: this.RecurrenceData
        }
        events.push(event);
      });
      //events = calendarEvents.data;

      events.sort(function (a, b) {
        var aDate = moment(a.EventDate);
        var bDate = moment(b.EventDate);
        return aDate.isAfter(bDate) ? 1 : (aDate.isBefore(bDate) ? -1 : 0);
      });

      return events;
    });

  }

  /*=====================================================
              Get current user details
    =======================================================*/
  /**
   * Get current user details
   */
  public async getCurrentUser(site: string = ""): Promise<IUser> {
    site = site == "" ? this.absoluteUrl : site;
    var url = site + "/_api/web/currentuser";
    var user: IUser = undefined;
    const result = await this.get(url);
    if (result.ok) {
      user = result.data;
      return user;
    } else {
      return user;
    }
  }



  /*=====================================================
              Find the component admin
    =======================================================*/
  /**
   * Find the component admin.
   * @param accessMatrixListName List name.
   * @param componentName List name.
   */

  public async isComponentAdmin(accessMatrixListName: string, componentName: string): Promise<boolean> {
    var url = this.absoluteUrl + `/_api/web/lists/getByTitle('${accessMatrixListName}')/items?$filter=Title eq '${componentName}'`;
    try {
      let userId = typeof this.pageContextInfo.userId == "string" ? parseInt(this.pageContextInfo.userId) : this.pageContextInfo.userId;
      const itemResult = await this.get(url);
      if (itemResult.ok && itemResult.data.value.length > 0) {
        //var currentUser = await this.getCurrentUser();
        if (itemResult.data.value[0].OwnerId.indexOf(userId) > -1) {
          return true;
        } else {
          let users = [];
          const userResults = itemResult.data.value[0].OwnerId.map(async (id: any) => {
            let groupUsers = await this.getGroupUsers(id, true);
            users = [...users, ...groupUsers];
          });
          await Promise.all(userResults);
          if (users.indexOf(userId) > -1) {
            return true;
          } else {
            return false;
          }
          // var userUrl = this.absoluteUrl + `/_api/web/sitegroups/getbyid(${itemResult.data.value[0].OwnerId})/users?$filter=Id eq ${this.pageContextInfo.userId}`;
          // const userResult = await this.get(userUrl);
          // if (userResult.ok && userResult.data.value.length > 0) {
          //   return true;
          // } else {
          //   console.log(userResult);
          //   return false;
          // }
        }
      } else {
        console.log(itemResult);
        return false;
      }
    }
    catch (err) {
      return false;
    }
  }

  /*=====================================================
              Get Site Groups
    =======================================================*/
  /**
   * get site groups.
   * @param isCurrentUserGroupsOnly get current user group only.
   * @param returnOnlyId to return only array of group id
   * @param site site name.
   */

  public async getSiteGroups(isCurrentUserGroupsOnly: boolean = false, returnOnlyId: boolean = false, site: string = "") {
    site = site == "" ? this.absoluteUrl : site;
    let url = site + "/_api/web/sitegroups";
    let expand = ["users"];
    const groups = await this.getPagedListItems("", [], "", "Title asc", expand, url);
    // const filteredGroup = groups.map((element) => {
    //   return {...element, Users: element.Users.filter((Users) => Users.Id === 1)}
    // })
    if (isCurrentUserGroupsOnly) {
      const filteredGroup = groups.filter(group => {
        return group.Users.results.filter(user => user.Id == this.pageContextInfo.userId).length > 0;
      });
      return returnOnlyId ? filteredGroup.map(group => group.Id) : filteredGroup;
    } else {
      return returnOnlyId ? groups.map(group => group.Id) : groups;
    }
  }

  /*=====================================================
              Get Group Users
    =======================================================*/
  /**
   * get site groups.
   * @param returnOnlyId to return only array of group id
   * @param site site name.
   */

  public async getGroupUsers(groupId: string | number, returnOnlyId: boolean = false, site: string = "") {
    site = site == "" ? this.absoluteUrl : site;
    let url = site + `/_api/web/sitegroups/getbyid(${groupId})/users?$top=4999`;
    let users = [];
    const groupResult = await this.get(url);
    if (groupResult.ok) {
      users = groupResult.data.value.length > 0 ? groupResult.data.value : [];
      return returnOnlyId ? users.map(user => user.Id) : users;
    } else {
      this.utils.error(groupResult, "getUserGroup");
      users.push(groupId);
      return users;
    }
  }

  /*=====================================================
                    Redirect
    =======================================================*/
  /**
   * Redirect.
   * @param url url.
   */

  public redirect(url: string = "") {
    if (this.utils.isNullOrEmpty(url)) {
      window.location.href = this.absoluteUrl;
    } else if (url.indexOf(this.absoluteUrl) > -1) {
      window.location.href = url;
    } else {
      window.location.href = this.absoluteUrl + url;
    }
  }

  /*=====================================================
                    Discussion Reply
    =======================================================*/
  /**
   * Redirect.
   * @param listTitle url.
   * @param messagePayload meta for reply
   * @example
   * var messagePayload = {
      '__metadata': { "type": "SP.Data.DistestListItem" },  //set DiscussionBoard entity type name
      'Body': "Final Test child",  //message Body
      'FileSystemObjectType': 0, //setto 0 to make sure Mesage Item
      'ContentTypeId': '0x0107008822E9328717EB48B3B665EE2266388E', //set Message content type
      'ParentItemID': 1  //set Discussion (topic) Id
    };
   */
  public createNewDiscussionReply(listTitle: string, messagePayload: any) {
    let __this = this;
    var topicUrl = null;
    var webUrl = this.absoluteUrl;
    return this.getParentTopic(webUrl, listTitle, messagePayload.ParentItemID)
      .then(function (result) {
        topicUrl = result.d.ServerRelativeUrl;
        return __this.createListItem(webUrl, listTitle, messagePayload);
      })
      .then(function (result) {
        var itemId = result.d.Id;
        return __this.moveListItem(webUrl, listTitle, itemId, topicUrl);
      });
  }
  private executeJson(options: any) {
    var headers = options.headers || {};
    var method = options.method || "GET";
    headers["Accept"] = "application/json;odata=verbose";
    if (options.method == "POST") {
      headers["X-RequestDigest"] = this.digest;
    }

    var ajaxOptions: any =
    {
      url: options.url,
      type: method,
      contentType: "application/json;odata=verbose",
      headers: headers
    };
    if ("data" in options) {
      ajaxOptions.data = JSON.stringify(options.data);
    }

    return $.ajax(ajaxOptions);
  }


  private createListItem(webUrl: string, listTitle: string, payload: any) {
    var url = webUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/items";
    return this.executeJson({
      "url": url,
      "method": 'POST',
      "data": payload
    });
  }

  private moveListItem(webUrl: string, listTitle: string, itemId: any, folderUrl: any) {
    let __this = this;
    var url = webUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/getItemById(" + itemId + ")?$select=FileDirRef,FileRef";
    return this.executeJson({
      "url": url
    })
      .then(function (result) {
        var fileUrl = result.d.FileRef;
        var fileDirRef = result.d.FileDirRef;
        var moveFileUrl = fileUrl.replace(fileDirRef, folderUrl);
        var url = webUrl + "/_api/web/getfilebyserverrelativeurl('" + fileUrl + "')/moveto(newurl='" + moveFileUrl + "',flags=1)";
        return __this.executeJson({
          "url": url,
          "method": 'POST'
        });
      });
  }


  private getParentTopic(webUrl: string, listTitle: string, itemId: any) {
    var url = webUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/getItemById(" + itemId + ")/Folder";
    return this.executeJson({
      "url": url,
    });
  }

  /*=====================================================
                   To Tag image to list item
   =======================================================*/
  /**
   * Redirect.
   * @param imageLibraryName name of library to upload image.
   * @param imageMeta metadata for the image
   * @param targetListName name of the list to tag the image
   * @param targetItemId item id
   * @param targetItemColumn hyperlink column name of the target list
   */
  public async tagImageToListItem(imageLibraryName: string, targetListName: string, targetItemId: string | number, targetItemColumn: string, elemId: string) {

    let serUrl = this.serverRelativeUrl == "/" ? "" : this.serverRelativeUrl;
    imageLibraryName = serUrl + imageLibraryName;
    let fi = document.getElementById(elemId) as HTMLInputElement;
    let fileName = fi.files.item(0).name;
    let folderIsExists = await this.isFolderExists(imageLibraryName);
    let fileServerUrl = imageLibraryName + "/" + fileName;
    if (!folderIsExists) {
      // create folder
      let _url = `${this.absoluteUrl}/_api/web/folders`;
      const meta = {
        "__metadata": {
          "type": "SP.Folder"
        },
        "ServerRelativeUrl": imageLibraryName
      }
      const res = await this.post(_url, meta);
    }

    const targetMeta = {};
    let url: IUrl = {
      Url: fileServerUrl,
      Description: fileName
    }
    targetMeta[targetItemColumn] = url;
    /*to check if the image is already uploaded for the item*/
    let fileIsExists = await this.isFileExists(fileServerUrl);
    if (!fileIsExists) {
      //Upload the given file
      await this.uploadFileToFolder(imageLibraryName, elemId);
      await this.updateItemById(targetListName, targetItemId, targetMeta);
    } else {
      //File already exists
      await this.updateItemById(targetListName, targetItemId, targetMeta);
    }
  }

  /*=====================================================
                   To download a file using blob
   =======================================================*/
  /**
   * Redirect.
   * @param fileServerRelativeUrl file server relative url.
   * @param fileName Name of the file
   * @param site site absolute url
   */
  public downloadFileUsingBlob(fileServerRelativeUrl: string, fileName: string, site: string = "") {
    let __this = this;
    site = site ? site : this.absoluteUrl;
    let filepath = `${site}/_api/Web/GetFileByServerRelativeUrl('${fileServerRelativeUrl}')/openbinarystream`;
    var dfd = $.Deferred();
    var xhr = new XMLHttpRequest();
    xhr.open("GET", filepath);
    xhr.responseType = "blob";
    //setting response-type header to be blob so that we get the file as blob
    xhr.onload = function () {
      //async call
      var blobobj = xhr.response;
      __this.utils.downloadBlob(blobobj, fileName);
      //window.navigator.msSaveBlob(blobobj, fileName);
      //save using msSaveBlob.
      dfd.resolve(true);
    }
    xhr.send();
    return dfd.promise()
  }

  /*=====================================================
                   To generate form number
   =======================================================*/
  /**
   * @param formNoConfigListName Form Number Configuration List Name.
   * @param formName Name of the form
   * @param site site absolute url
   */
  public async generateFormNo(formNoConfigListName: string, formName: string, site: string = "") {
    site = site ? site : this.absoluteUrl;
    let formNo = "";
    let formNoListUrl = site + `/_api/lists/getbytitle('${formNoConfigListName}')/items?$filter=Title eq '${formName}'`;
    const formNoResult = await this.get(formNoListUrl);
    if (formNoResult.ok && formNoResult.data.value.length > 0) {
      let item: IFormNoConfiguration = formNoResult.data.value[0];
      const today = new Date();
      let month = (today.getMonth() + 1);
      let year = (month >= 4) ? today.getFullYear() : (today.getFullYear() - 1); // 4 - April
      formNo += item.Prefix1 ? item.Prefix1 + "/" : "";
      formNo += item.Prefix2 ? item.Prefix2 + "/" : "";
      formNo += `FY${year}/`;
      formNo += this.utils.pad(Math.floor(item.UpcomingFormNo).toString(), item.NumberLength, "0");
      let meta: IFormNoConfiguration = {
        CurrentFormNo: item.UpcomingFormNo
      }
      await this.updateItemById(formNoConfigListName, item.Id, meta);
    }
    return formNo;
  }

  /*=====================================================
                   To get email template
   =======================================================*/
  /**
   * Redirect.
   * @param templateName name of the template.
   * @param listName Name of the list
   * @param placeholderDetails metadata for placeholder
   */
  public async getEmailTemplate(listName: string, templateName: string, placeholderDetails: any) {
    let templateUrl = this.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/items?$filter=Title eq '${templateName}'`;
    const result = await this.get(templateUrl);
    if (result.ok && result.data.value.length > 0) {
      let item: IEmailTemplate = result.data.value[0];
      var sentence = item.Body;//"My name is $Name$ and age is $Age$"
      var placeholders = sentence.match(/\$(.*?)\$/g); // '{' = &#123; '}' = &#125;
      placeholders.forEach(function (placeholder) {
        //Placeholder - $Name$
        var phText = placeholder.substring(1, placeholder.length - 1);
        //phText = Name

        if (placeholderDetails[phText]) {
          sentence = sentence.replace(placeholder, placeholderDetails[phText])
        } else {
          sentence = sentence.replace(placeholder, "N/A")
        }
      });
      console.log(sentence);
      return sentence;
    } else {
      if (!result.ok) {
        this.utils.error(result, "Email template", false);
      }
      return "";
    }
  }

  public async getDelegatedDetails(listName: string, userId: number) {
    let currentDateIso = new Date().toISOString().split("T")[0] + "T00:00:00Z";
    let delegateeId = 0;
    let delegateeDetails = undefined;
    let getRequestUrl = this.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/items?$select=Id,DelegateeId,Delegatee/Title&$expand=Delegatee&$filter=UserId eq '${userId}' and IsEnabled eq 1 and StartDate le '${currentDateIso}' and EndDate ge '${currentDateIso}'`;
    let responseData = await this.get(getRequestUrl);
    if (responseData.ok) {
      let delegateeIdLevel1 = (responseData.data.value.length > 0 ? responseData.data.value[0].DelegateeId : 0);
      if (delegateeIdLevel1 > 0) {
        let getRequestUrlLevel2 = this.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/items?$select=Id,DelegateeId,Delegatee/Title&$expand=Delegatee&$filter=UserId eq '${delegateeIdLevel1}' and IsEnabled eq 1 and StartDate le '${currentDateIso}' and EndDate ge '${currentDateIso}'`;
        let responseDataLevel2 = await this.get(getRequestUrlLevel2);
        let delegateeIdLevel2 = (responseDataLevel2.data.value.length > 0 ? responseDataLevel2.data.value[0].DelegateeId : 0);
        if (delegateeIdLevel2 > 0) {  // userId delegatee has another delegatee case
          delegateeId = userId;
          delegateeDetails = responseData.data.value[0];
        } else {
          delegateeId = delegateeIdLevel1;
          delegateeDetails = responseData.data.value[0];
        }
      }
      return delegateeDetails;
    } else {
      this.utils.error(responseData, "Out of Office", false);
      return -1;
    }
  }

  public async getDelegatedUserId(listName: string, userId: number) {
    let currentDateIso = new Date().toISOString().split("T")[0] + "T00:00:00Z";
    let delegateeId = 0;
    let getRequestUrl = this.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/items?$select=Id,DelegateeId&$filter=UserId eq '${userId}' and IsEnabled eq 1 and StartDate le '${currentDateIso}' and EndDate ge '${currentDateIso}'`;
    let responseData = await this.get(getRequestUrl);
    if (responseData.ok) {
      let delegateeIdLevel1 = (responseData.data.value.length > 0 ? responseData.data.value[0].DelegateeId : 0);
      if (delegateeIdLevel1 > 0) {
        let getRequestUrlLevel2 = this.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/items?$select=Id,DelegateeId&$filter=UserId eq '${delegateeIdLevel1}' and IsEnabled eq 1 and StartDate le '${currentDateIso}' and EndDate ge '${currentDateIso}'`;
        let responseDataLevel2 = await this.get(getRequestUrlLevel2);
        let delegateeIdLevel2 = (responseDataLevel2.data.value.length > 0 ? responseDataLevel2.data.value[0].DelegateeId : 0);
        if (delegateeIdLevel2 > 0) {  // userId delegatee has another delegatee case
          delegateeId = userId;
        } else {
          delegateeId = delegateeIdLevel1;
        }
      }
      return delegateeId;
    } else {
      this.utils.error(responseData, "Out of Office", false);
      return -1;
    }
  }

  public async getMyGraphAPIData(props): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      props.context.msGraphClientFactory.getClient()
        .then((client: MSGraphClient) => {
          // get information about the current user from the Microsoft Graph
          return client.api('/me').version('beta').get((error, response: any, rawResponse?: any) => {
            if (error) {
              console.error(error);
              reject(error);
            }
            resolve(response);
          });
        });
    });
  }

}