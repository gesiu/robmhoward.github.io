Word.run(function (ctx) {
var paras = ctx.document.body.paragraphs;
ctx.load(paras);

ctx.sync()
    .then(function () {
        paras.items[0].insertInlinePictureFromBase64("iVBORw0KGgoAAAANSUhEUgAAAIAAAACABAMAAAAxEHz4AAAAJFBMVEX///9GRkZGRkZGRkZGRkZGRkZGRkZGRkYBpO9/ugDyUCL/uQGm4PjWAAAACHRSTlMBCQ0RFRknMx7uViEAAAB3SURBVGje7dcxCYBQGEXhi6izYBHB0RIiiAXkzW5iAMEKFnCwguVscJd/ecM5Ab79SNHK5FqlZXeNql/XIx23awMAAAAAAAAAAAAAAAAAyBwIvzNJxeyapLZ3Naou1ykNn6sDAAAAAAAAAAAAAAAAAMgcCL9ztB/UhshWs1l/WAAAAABJRU5ErkJggg==", Word.InsertLocation.end);
        var pics = ctx.document.body.inlinePictures
        ctx.load(pics);

        ctx.sync()
            .then(function () {
                console.log("Picture Count: " + pics.items.length);
                console.log("Success");
            });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
})