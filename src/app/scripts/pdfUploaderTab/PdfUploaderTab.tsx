import * as React from "react";
import { Provider, Flex, Loader, FilesPdfColoredIcon, RedoIcon } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwt_decode from "jwt-decode";
import Axios from "axios";
import Utilities from "../../api/Utilities";

/**
 * Implementation of the PDF Uploader content page
 */
export const PdfUploaderTab = () => {
    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [token, setToken] = useState<string>();
    const [highlight, setHighlight] = useState<boolean>(false);
    const [siteDomain, setSiteDomain] = useState<string>();
    const [sitePath, setSitePath] = useState<string>();
    const [channelName, setChannelName] = useState<string>();
    const [status, setStatus] = useState<string>();
    const [uploadUrl, setUploadUrl] = useState<string>();
    
    useEffect(() => {
        if (inTeams === true) {

            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    setToken(token);
                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [`api://${process.env.HOSTNAME}/${process.env.PDFUPLOADER_APP_ID}` as string]
            });
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.entityId);
            setSiteDomain(context.teamSiteDomain);
            setSitePath(context.teamSitePath);
            setChannelName(context.channelName);
        }
    }, [context]);

    const allowDrop = (event) => {
            event.preventDefault();
            event.stopPropagation();
            event.dataTransfer.dropEffect = 'copy';
    };
    const enableHighlight = (event) => {
            allowDrop(event);
            setHighlight(true);
    };
    const disableHighlight = (event) => {
            allowDrop(event);
            setHighlight(false);
    };
    const dropFile = (event) => {
        allowDrop(event);
        const dt = event.dataTransfer;
        const files =  Array.prototype.slice.call(dt.files); 
        files.forEach(fileToUpload => {
            if (Utilities.validFileExtension(fileToUpload.name)) {
                uploadFile(fileToUpload);
            }
        });
    };

    const uploadFile = (fileToUpload: File) => {
        setStatus('running');
        const formData = new FormData();
        formData.append('file', fileToUpload);
        formData.append('domain', siteDomain!);
        formData.append('sitepath', sitePath!);
        formData.append('channelname', channelName!);
        Axios.post(`https://${process.env.HOSTNAME}/api/upload`, formData, {
                                    headers: {
                                        'Authorization': `Bearer ${token}`,
                                        'content-type': 'multipart/form-data'
                                    }
                                }).then(result => {
                                    console.log(result);
                                    setStatus('uploaded');
                                    setUploadUrl(result.data);
                                });
    };

    const reset = () => {
        setStatus('');
        setUploadUrl('');
    };
        
    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <div className='dropZoneBG'>
                  Drag your file here:
                  <div className={`${highlight ? 'dropZone dropZoneHighlight':'dropZone'}`}
                                        onDragEnter={enableHighlight} 
                                        onDragLeave={disableHighlight} 
                                        onDragOver={allowDrop} 
                                        onDrop={dropFile}>
                    {status !== 'running' && status !== 'uploaded' &&
                            <div className='pdfLogo'>
                                <FilesPdfColoredIcon size="largest" bordered />
                            </div>}
                    {status === 'running' &&
                            <div className='loader'>
                                <Loader label="Upload and conversion running..." size="large" labelPosition="below" inline />
                            </div>}
                    {status === 'uploaded' && 
                            <div className='result'>File uploaded to target and available <a href={uploadUrl}>here.</a>
                            <RedoIcon size="medium" bordered onClick={reset} title="Reset" /></div>}

                  </div>
                </div>
            </Flex>
        </Provider>
    );
};
