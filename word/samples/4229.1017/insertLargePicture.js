
var t0 = performance.now();
var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);
ctx.references.add(paras)

ctx.executeAsync()
    .then(function() {
        for (var i = 0; i < 1; i++) {
            paras.items[0].insertInlinePictureFromBase64(base64_img, Word.InsertLocation.end);
        }
        ctx.references.remove(paras);
    })
    .then(ctx.executeAsync)
    .then(function() {
        var t1 = performance.now();
        console.log("Elapsed time = " + (t1 - t0)/1000);
    })
    .catch(function(error) {
        console.log(JSON.stringify(error));
    });