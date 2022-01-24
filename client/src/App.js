import React, { useRef, useEffect } from "react";
import WebViewer from "@pdftron/webviewer";
import axios from "axios";
import "./App.css";
import qs from "qs";

const body = {
  grant_type: process.env.REACT_APP_GRANT_TYPE,
  client_id: process.env.REACT_APP_CLIENT_ID,
  client_secret: process.env.REACT_APP_CLIENT_SECRET,
  resource: process.env.REACT_APP_RESOURCE,
};

const getDownloadFileName = (name, extension = ".pdf") => {
  if (name.slice(-extension.length).toLowerCase() !== extension) {
    name += extension;
  }
  return name;
};

const getToken = () => {
  return axios({
    method: "post",
    url: `/${process.env.REACT_APP_TENANT_ID}/tokens/OAuth/2`,
    headers: {
      Accept: "application/json;odate=verbose",
      "Content-Type": "application/x-www-form-urlencoded",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Credentials": true,
    },
    data: qs.stringify(body),
  }).then(({ data }) => {
    const token = data.access_token;
    return token;
  });
};

const uploadFile = ({ file, foldername, downloadFileName, token }) => {
  console.log(file, foldername, downloadFileName, token);
  return axios
    .post(
      `${process.env.REACT_APP_ABSOLUTE_URL}/_api/web/GetFolderByServerRelativeUrl('${foldername}')/Files/add(url='${downloadFileName}',overwrite=true)`,
      file,
      {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      }
    )
    .then(({ data }) => {
      const { ServerRelativeUrl } = data;
      return axios
        .get(
          `${process.env.REACT_APP_ABSOLUTE_URL}/_api/Web/GetFileByServerRelativePath(decodedurl='${ServerRelativeUrl}')/ListItemAllFields`,
          {
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
            },
          }
        )
        .then(({ data }) => {
          return {
            id: data.d.ID,
            token,
            guid: data.d.GUID,
          };
        });
    })
    .catch((err) => console.log(err));
};

const App = () => {
  const viewer = useRef(null);
  useEffect(() => {
    WebViewer(
      {
        path: "/webviewer/lib",
      },
      viewer.current
    ).then((instance) => {
      const { docViewer, annotManager, UI } = instance;
      const urlParams = new URLSearchParams(window.location.search);
      const uniqueId = urlParams.get("uniqueId");
      const tempAuth = urlParams.get("tempAuth");
      const filename = urlParams.get("filename");
      const foldername = urlParams.get("foldername");
      const username = urlParams.get("username") || "Guest";
      const email = urlParams.get("email");
      tempAuth &&
        instance.loadDocument(
          `${process.env.REACT_APP_ABSOLUTE_URL}/_layouts/15/download.aspx?UniqueId=${uniqueId}&Translate=false&tempauth=${tempAuth}&ApiVersion=2.0`,
          { filename }
        );
      annotManager.setCurrentUser(username);
      annotManager.setIsAdminUser(true);
      foldername &&
        instance.UI.settingsMenuOverlay.add(
          {
            type: "actionButton",
            className: "row",
            img: '<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="24px" height="24px" viewBox="0 0 512 512" style="enable-background:new 0 0 512 512;" xml:space="preserve"> <path d="M412.907,214.08C398.4,140.693,333.653,85.333,256,85.333c-61.653,0-115.093,34.987-141.867,86.08 C50.027,178.347,0,232.64,0,298.667c0,70.72,57.28,128,128,128h277.333C464.213,426.667,512,378.88,512,320 C512,263.68,468.16,218.027,412.907,214.08z M298.667,277.333v85.333h-85.333v-85.333h-64L256,170.667l106.667,106.667H298.667z" /> </svg>',
            onClick: () => {
              instance.openElements([modal.dataElement]);
            },
            dataElement: "uploadFile",
            label: "Upload file",
          },
          "themeChangeButton"
        );
      UI.enableFeatures([UI.Feature.FilePicker]);
      var modal = {
        dataElement: "meanwhileInFinlandModal",
        render: function renderCustomModal() {
          const doc = docViewer.getDocument();
          var div = document.createElement("div");
          const para = document.createElement("p");
          const node = document.createTextNode(
            "Enter your custom file name, please!"
          );
          para.appendChild(node);
          div.appendChild(para);
          var input = document.createElement("INPUT");
          input.value = doc.getFilename();
          input.setAttribute("type", "text");
          input.setAttribute("id", "myInput");
          var button = document.createElement("BUTTON");
          div.appendChild(input);
          button.setAttribute("content", "confirm");
          button.textContent = "Save";
          button.onclick = async () => {
            const doc = docViewer.getDocument();
            const downloadFileName = getDownloadFileName(input.value);
            const xfdfString = await annotManager.exportAnnotations();
            const data = await doc.getFileData({
              // saves the document with annotations in it
              xfdfString,
              downloadType: "pdf",
            });
            const arr = new Uint8Array(data);
            const file = new File([arr], downloadFileName, {
              type: "application/pdf",
            });
            getToken()
              .then((token) =>
                uploadFile({
                  file,
                  foldername,
                  downloadFileName,
                  token,
                })
              )
              .then(({ token, id }) => {
                const formValues = [
                  {
                    FieldName: "Editor",
                    FieldValue: `[{'Key':'i:0#.f|membership|${email}'}]`,
                  },
                ];
                const title =
                  foldername === "Shared Documents" ? "Documents" : foldername;
                axios.post(
                  `${
                    process.env.REACT_APP_ABSOLUTE_URL
                  }/_api/Web/Lists/getByTitle('${
                    title.split("/")[0]
                  }')/Items(${id})/ValidateUpdateListItem`,
                  JSON.stringify({
                    formValues,
                    bNewDocumentUpdate: true,
                  }),
                  {
                    headers: {
                      Authorization: `Bearer ${token}`,
                      Accept: "application/json;odata=verbose",
                      "If-Match": "*",
                      "Content-Type": "application/json;odata=verbose",
                    },
                  }
                );
              });
            instance.closeElements([modal.dataElement]);
          };
          div.appendChild(button);
          div.style.color = "black";
          div.style.backgroundColor = "white";
          div.style.fontSize = "16px";
          div.style.padding = "20px 40px";
          div.style.borderRadius = "5px";
          return div;
        },
      };
      instance.setCustomModal(modal);
    });
  }, []);

  return (
    <div className="App">
      <div className="webviewer" ref={viewer}></div>
    </div>
  );
};

export default App;
