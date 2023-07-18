// script.js
// const notificationSource = new EventSource('/notify');
const notificationTextarea = document.getElementById('notification-textarea');
let notificationSource;

function startNotifications() {
    // Check if the SSE connection is already established
    if (notificationSource) {
        return;
    }
    console.log('Start Notice');
    notificationSource = new EventSource('/notify');

    notificationSource.onmessage = function (event) {
        const data = event.data;
        if (data) {
            console.log(data)
            // Append the new notification to the textarea
            notificationTextarea.value += data + '\n';
            // Scroll to the bottom of the textarea
            notificationTextarea.scrollTop = notificationTextarea.scrollHeight;
        }
    };

    notificationSource.onerror = function (event) {
        console.error('Error occurred:', event);
        notificationSource.close();
    };
};

function closeNotifications() {
    if (notificationSource) {
        console.log('Close Notice');
        notificationSource.close();
        notificationSource = null;
    }
}