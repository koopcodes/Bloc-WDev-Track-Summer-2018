/*
 * Programming Quiz: Facebook Friends (7-5)
 */

// your code goes here
var facebookProfile = {
    name: "Brent Kupras",
    friends: 589,
    messages: ["I love coding.", "I do too.", "Me three!"],
    postMessage: function postMessage(message){
        facebookProfile.messages.push(message);
    },
    deleteMessage: function deleteMessage(index) {
        facebookProfile.messages.splice(index, 1);
    },
    addFriend: function addFriend() {
        facebookProfile.friends = facebookProfile.friends + 1;
    },
    removeFriend: function removeFriend() {
        facebookProfile.friends = facebookProfile.friends- 1;
    },
};