var wordSamplesApp = angular.module("wordSamplesApp", ['ngRoute']);
var insideOffice = false;
var debugOption = false;
var officeVersion;

var logComment = function (message) {
    var span = document.createElement('span');
    span.className = 'message-text';
    span.innerHTML = message + '<br/>';
    $('#console').append(span);
}

var logDebug = function (message) {
    if (debugOption) {
        var span = document.createElement('span');
        span.className = 'debug-text';
        span.innerHTML = message + '<br/>';
        $('#console').append(span);
    }
}

function initializeMonacoEditor() {
    $('#TxtRichApiScript').empty();
    
    // Update to full path if word is not at the root folder
    MonacoEditorIntegration.initializeJsEditor('TxtRichApiScript', [
            "/word/script/" + officeVersion + "/EditorIntelliSense/WordLatest.txt",
            "/word/script/" + officeVersion + "/EditorIntelliSense/Office.Runtime.txt",
            "/word/script/EditorIntelliSense/Helpers.txt",
            "/word/script/EditorIntelliSense/jquery.txt",
    ]);
    
    var ua = navigator.userAgent;    
    var ipad = ua.indexOf("iPad");    
    if (ipad >= 0)
    {
         $("TxtRichApiScript").addClass("script_box_ipad");
    }
}

function createJSFile (filename) {
    var fileRef = document.createElement('script');
    fileRef.setAttribute("type","text/javascript");
    fileRef.setAttribute("src", filename);

    var head = document.getElementsByTagName('head')[0];
    head.appendChild(fileRef);
}

// Unused function
function replaceJSFile (oldFilename, newFilename) {
    var allScripts = document.getElementsByTagName("script");

    for (var i = allScripts.length; i >= 0; i--) {
        if (allScripts[i]) {
            var sourceFilename = allScripts[i].getAttribute("src");
            if (sourceFilename != null && sourceFilename.indexOf(oldFilename) != -1) {
                var newElement = createJSFile(newFilename);
                allScripts[i].parentNode.replaceChild(newElement, allScripts[i]);
                return;
            }
        }
    }
}

Office.initialize = function (reason) {
    insideOffice = true;

    // Override window.console to log framework debug info
    window.console.log = function (message) {
        logDebug(message);
    };

    logDebug('Initialized!');
};

wordSamplesApp.config(['$routeProvider', function ($routeProvider) {
    $routeProvider
        .when('/samples',
            {
                controller: 'SamplesController',
                templateUrl: 'partials/samples.html'
            })
        .otherwise({ redirectTo: '/samples' });
}]);

wordSamplesApp.factory("wordSamplesFactory", ['$http', function ($http) {
    var factory = {};

    factory.getSamples = function () {
        return $http.get('samples/' + officeVersion + '/samples.json');
    };

    factory.getSampleCode = function (filename) {
        return $http.get('samples/' + officeVersion + '/' + filename);
    };

    return factory;
}]);

wordSamplesApp.controller("SamplesController", function ($scope, wordSamplesFactory) {
    $scope.samples = [{ name: "Loading..." }];
    $scope.builds = [{name: "6409.1000"}, { name: "4229.1017"}, { name: "wac"}];
    $scope.selectedSample = { description: "No sample loaded" };
    $scope.selectedBuild = $scope.builds[0];
    $scope.debugOption = { value: false };
    $scope.insideOffice = insideOffice;

    $scope.switchOfficeVersion = function() {
        officeVersion = $scope.selectedBuild.name;

        // Reload Monaco Editor
        initializeMonacoEditor();

        // The new office.js (referenced in app.html) includes everything, 
		// so loading extra JavaScript is not necessary.
        // createJSFile('script/' + officeVersion + '/Office.Runtime.js');
        // createJSFile('script/' + officeVersion + '/Word.js');
        
        // Reload samples
        wordSamplesFactory.getSamples().then(function (response) {
            $scope.samples = response.data.values;
            $scope.groups = response.data.groups;
        });
    }
    
    $scope.switchOfficeVersion();

    MonacoEditorIntegration.setDirty = function () {
        if ($scope.selectedSample.code) {
            $scope.selectedSample = { description: $scope.selectedSample.description + " (modified)" };
            $scope.$apply();
        }
    }

    $scope.loadSampleCode = function () {
        console.log("loadSampleCode called");
        appInsights.trackEvent("SampleLoaded", { name: $scope.selectedSample.name });
        wordSamplesFactory.getSampleCode($scope.selectedSample.filename).then(function (response) {
            $scope.selectedSample.code = response.data;
            $scope.insideOffice = insideOffice;
            MonacoEditorIntegration.setJavaScriptText($scope.selectedSample.code);
        });
    };

    $scope.runSelectedSample = function () {
        var script = MonacoEditorIntegration.getJavaScriptToRun().replace(/console\.log/g, "logComment");
        script = "try {" + script + "} catch (e) { logComment(\"Exception: \" + e.message ? e.message : e);}";

        logComment("====="); // Add separators between executions
        try {
            eval(script);
        } 
        catch (ex) {
            logDebug(ex.name + ": " + ex.message)
        }
    }

    $scope.toggleDebugOption = function () {
        debugOption = $scope.debugOption.value;
    }

    $scope.clearLog = function () {
        $('#console').empty();
    }
});
