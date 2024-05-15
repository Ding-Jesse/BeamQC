// script.js
// const notificationSource = new EventSource('/notify');
const notificationTextarea = document.getElementById('notification-textarea');
let notificationSource;

function startNotifications(url) {
    // Check if the SSE connection is already established
    if (notificationSource) {
        console.log("An EventSource connection is already active.");
        return;
    }
    console.log('Start Notice');
    notificationSource = new EventSource(url);

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

    window.addEventListener('unload', () => {
        if (notificationSource) {
            eventSource.close();
        }
    });
};

function closeNotifications() {
    if (notificationSource) {
        console.log('Close Notice');
        notificationSource.close();
        notificationSource = null;
    }
}
function testNotification() {
    const notificationTextarea = document.getElementById('notification-textarea');
    var eventSource = new EventSource('/stream-logs');
    eventSource.onmessage = function (event) {
        console.log('New log entry:', event.data);
        // Append the new notification to the textarea
        notificationTextarea.value += event.data + '\n';
        // Scroll to the bottom of the textarea
        notificationTextarea.scrollTop = notificationTextarea.scrollHeight;
        if (event.data.includes("EOF")) {
            eventSource.close();
            notificationTextarea.value += 'Finish' + '\n';
            console.log('Finish');
        }
    }
    // Handle any errors that occur
    eventSource.onerror = function (error) {
        console.error('EventSource failed:', error);
        eventSource.close();  // Close the connection on error
    };
    // To close the EventSource connection when the user navigates away
    window.onbeforeunload = function () {
        eventSource.close();
    };
}