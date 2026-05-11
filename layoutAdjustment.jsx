
//Master Page text frames not adjusted automatically when change the margin

var myDocument = app.activeDocument;

try {
	myDocument.adjustLayoutPreferences.enableAdjustLayout = true;
	myDocument.adjustLayoutPreferences.enableAutoAdjustMargins = true;
	myDocument.adjustLayoutPreferences.allowLockedObjectsToAdjust = true;
} catch (e) {
	alert(e); exit(0);
}

var myMaster = myDocument.masterSpreads;

for (var k = 0; k < myMaster.length; k++) {
    var pages = myMaster[k].pages;
    for (var p = 0; p < pages.length; p++) {
        var myPage = pages[p];
        if (myPage.side == PageSideOptions.LEFT_HAND) {
            if (myPage.marginPreferences.left < 63)
                myPage.marginPreferences.left = 63;
            if (myPage.marginPreferences.right < 36)
                myPage.marginPreferences.right = 36;

        } else if (myPage.side == PageSideOptions.RIGHT_HAND) {
            if (myPage.marginPreferences.left < 63)
                myPage.marginPreferences.left = 63;
            if (myPage.marginPreferences.right < 36)
                myPage.marginPreferences.right = 36;
        }

        var mp = myPage.marginPreferences;
        var pb = myPage.bounds; // [top, left, bottom, right]

        var frameTop = pb[0] + mp.top;
        var frameLeft = pb[1] + mp.left;
        var frameBottom = pb[2] - mp.bottom;
        var frameRight = pb[3] - mp.right;

        var frames = myPage.textFrames;
        for (var f = 0; f < frames.length; f++) {
            frames[f].geometricBounds = [frameTop, frameLeft, frameBottom, frameRight];
        }
    }
}
