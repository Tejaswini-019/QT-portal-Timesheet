import SPService from "./SPService";

var _spPageContextInfo: any;  
var SP: any;    
export class FileUploadService {  
    
    public siteUrl: string = "http://sps2016-hta";  
    public digest: string = "";
    public siteRelativeUrl: string = "/" ;//_spPageContextInfo.webServerRelativeUrl != "/" ? _spPageContextInfo.webServerRelativeUrl : "";  
    public fileUpload(file: any, documentLibrary: string, fileName: string, digest: string) {
        this.digest = digest;  
        return new Promise((resolve, reject) => {  
            this.createDummyFile(fileName, documentLibrary).then(result => {  
                let fr = new FileReader();  
                let offset = 0;  
                // the total file size in bytes...  
                let total = file.size;  
                // 1MB Chunks as represented in bytes (if the file is less than a MB, seperate it into two chunks of 80% and 20% the size)...  
                let length = parseInt("1000000") > total ? Math.round(total * 0.8) : parseInt("1000000");  
                let chunks = [];  
                //reads in the file using the fileReader HTML5 API (as an ArrayBuffer) - readAsBinaryString is not available in IE!  
                fr.readAsArrayBuffer(file);  
                fr.onload = (evt: any) => {  
                    while (offset < total) {  
                        //if we are dealing with the final chunk, we need to know...  
                        if (offset + length > total) {  
                            length = total - offset;  
                        }  
                        //work out the chunks that need to be processed and the associated REST method (start, continue or finish)  
                        chunks.push({  
                            offset,  
                            length,  
                            method: this.getUploadMethod(offset, length, total)  
                        });  
                        offset += length;  
                    }  
                    //each chunk is worth a percentage of the total size of the file...  
                    const chunkPercentage = (total / chunks.length) / total * 100;  
                    if (chunks.length > 0) {  
                        //the unique guid identifier to be used throughout the upload session  
                        const id = this.guid();  
                        //Start the upload - send the data to S  
                        this.uploadFile(evt.target.result, id, documentLibrary, fileName, chunks, 0, 0, chunkPercentage, resolve, reject);  
                    }  
                };  
            })  
        });  
    }  
    createDummyFile(fileName, libraryName) {  
        return new Promise((resolve, reject) => {  
            // Construct the endpoint - The GetList method is available for SharePoint Online only.  
            var serverRelativeUrlToFolder = "decodedurl='" + this.siteRelativeUrl + "/" + libraryName + "'";  
            var endpoint = this.siteUrl + "/_api/Web/GetFolderByServerRelativeUrl('/VideoMediaGallery')/files" + "/add(overwrite=true, url='" + fileName + "')"  
            const headers = {  
                "accept": "application/json;odata=verbose"  
            };  
            this.executeAsync(endpoint, this.convertDataBinaryString(2), headers).then(file => resolve(true)).catch(err => reject(err));  
        });  
    }  
    // Base64 - this method converts the blob arrayBuffer into a binary string to send in the REST request  
    convertDataBinaryString(data) {  
        let fileData = '';  
        let byteArray = new Uint8Array(data);  
        for (var i = 0; i < byteArray.byteLength; i++) {  
            fileData += String.fromCharCode(byteArray[i]);  
        }  
        return fileData;  
    }  
    executeAsync(endPointUrl, data, requestHeaders) {  
        return new Promise((resolve, reject) => {  
            // var scriptbase = "http://sps2016-hta/_layouts/15/";
            // $.getScript(scriptbase + "SP.RequestExecutor.js", function() {
            //     // using a utils function we would get the APP WEB url value and pass it into the constructor...  
            //     let executor = new SP.RequestExecutor(this.siteUrl);  
            //     // Send the request.  
            //     executor.executeAsync({  
            //         url: endPointUrl,  
            //         method: "POST",  
            //         body: data,  
            //         binaryStringRequestBody: true,  
            //         headers: requestHeaders,  
            //         success: offset => resolve(offset),  
            //         error: err => reject(err.responseText)  
            //     });
            // }); 
            $.ajax({
                url: endPointUrl,
                type: "post",
                data: data,
                processData: false,
                headers: {
                  accept: "application/json;odata=verbose",
                  "X-RequestDigest": this.digest,
                  "content-length": data.byteLength,
                },
                success: offset => resolve(offset),  
                error: err => reject(err.responseText)  
            });
        });  
    }  
    //this method sets up the REST request and then sends the chunk of file along with the unique indentifier (uploadId)  
    uploadFileChunk(id, libraryPath, fileName, chunk, data, byteOffset) {  
        return new Promise((resolve, reject) => {  
            let offset = chunk.offset === 0 ? '' : ',fileOffset=' + chunk.offset;  
            //parameterising the components of this endpoint avoids the max url length problem in SP (Querystring parameters are not included in this length)  
            let endpoint = this.siteUrl + "/_api/web/getfilebyserverrelativeurl('" + this.siteRelativeUrl + "/" + libraryPath + "/" + fileName + "')/" + chunk.method + "(uploadId=guid'" + id + "'" + offset + ")";  
            const headers = {  
                "Accept": "application/json; odata=verbose",  
                "Content-Type": "application/octet-stream"  
            };  
            this.executeAsync(endpoint, data, headers).then(offset => resolve(offset)).catch(err => reject(err));  
        });  
    }  
    //the primary method that resursively calls to get the chunks and upload them to the library (to make the complete file)  
    uploadFile(result, id, libraryPath, fileName, chunks, index, byteOffset, chunkPercentage, resolve, reject) {  
        //we slice the file blob into the chunk we need to send in this request (byteOffset tells us the start position)  
        const data = this.convertFileToBlobChunks(result, byteOffset, chunks[index]);  
        //upload the chunk to the server using REST, using the unique upload guid as the identifier  
        this.uploadFileChunk(id, libraryPath, fileName, chunks[index], data, byteOffset).then(value => {  
            const isFinished = index === chunks.length - 1;  
            index += 1;  
            const percentageComplete = isFinished ? 100 : Math.round((index * chunkPercentage));  
            //More chunks to process before the file is finished, continue  
            if (index < chunks.length) {  
                this.uploadFile(result, id, libraryPath, fileName, chunks, index, byteOffset, chunkPercentage, resolve, reject);  
            } else {  
                resolve(value);  
            }  
        }).catch(err => {  
            console.log('Error in uploadFileChunk! ' + err);  
            reject(err);  
        });  
    }  
    //Helper method - depending on what chunk of data we are dealing with, we need to use the correct REST method...  
    getUploadMethod(offset, length, total) {  
        if (offset + length + 1 > total) {  
            return 'finishupload';  
        } else if (offset === 0) {  
            return 'startupload';  
        } else if (offset < total) {  
            return 'continueupload';  
        }  
        return null;  
    }  
    //this method slices the blob array buffer to the appropriate chunk and then calls off to get the BinaryString of that chunk  
    convertFileToBlobChunks(result, byteOffset, chunkInfo) {  
        let arrayBuffer = result.slice(chunkInfo.offset, chunkInfo.offset + chunkInfo.length);  
        return this.convertDataBinaryString(arrayBuffer);  
    }  
    guid() {  
        function s4() {  
            return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1);  
        }  
        return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();  
    }  
}  