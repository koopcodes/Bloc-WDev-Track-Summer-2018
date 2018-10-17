// Brent Kupras
// Hack Reactor Application

var myArray = [];
myArray = ["Brent Kupras", "thinking.bear@gmail.com"];

function cutName (name) {
     var nameSplit = [];
     nameSplit = name.split(" ");
     return nameSplit;
}

var myInfo = {};
myInfo.fullName = cutName(myArray[0]);
myInfo.skype = myArray[1];
myInfo.gitHub = "UrsaCogitabundus";
