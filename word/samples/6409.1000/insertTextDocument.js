var ctx = new Word.RequestContext();

var textSample =
    "Hello, world! This is an example of the insert text method. This is a method which allows users to insert text at the end of the document. It also can insert text into a relative location.";

ctx.document.body.insertParagraph(textSample, Word.InsertLocation.end);

ctx.sync()
    .then(function () {
         console.log("Success");
     })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
