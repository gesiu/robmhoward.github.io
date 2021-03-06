Word.run(function (ctx) {

var mySections = ctx.document.sections;
ctx.load(mySections);
ctx.references.add(mySections);

ctx.sync()
    .then(function () {
        var paras = mySections.items[0].body.paragraphs;
        ctx.load(paras);

        ctx.sync()
            .then(function () {
                console.log("Number of paragraphs in section: " + paras.items.length);
            });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
})