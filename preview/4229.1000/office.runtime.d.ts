declare module OfficeExtension {
    class Action {
        private m_actionInfo;
        private m_isWriteOperation;
        constructor(actionInfo: ActionInfo, isWriteOperation: boolean);
        actionInfo: ActionInfo;
        isWriteOperation: boolean;
    }
}
declare module OfficeExtension {
    class ActionFactory {
        static createSetPropertyAction(context: ClientRequestContext, parent: ClientObject, propertyName: string, value: any): Action;
        static createMethodAction(context: ClientRequestContext, parent: ClientObject, methodName: string, operationType: OperationType, args: Array<any>): Action;
        static createQueryAction(context: ClientRequestContext, parent: ClientObject, queryOption: QueryInfo): Action;
        static createInstantiateAction(context: ClientRequestContext, obj: ClientObject): Action;
        static createTraceAction(context: ClientRequestContext, message: string): Action;
    }
}
declare module OfficeExtension {
    class ClientObject implements IResultHandler {
        private m_context;
        private m_objectPath;
        constructor(context: ClientRequestContext, objectPath: ObjectPath);
        context: ClientRequestContext;
        _objectPath: ObjectPath;
        _handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    class ClientRequest {
        private m_flags;
        private m_context;
        private m_actions;
        private m_actionResultHandler;
        private m_referencedObjectPaths;
        private m_traceInfos;
        constructor(context: ClientRequestContext);
        flags: ClientRequestFlags;
        traceInfos: {
            [index: number]: string;
        };
        addAction(action: Action): void;
        addTrace(actionId: number, message: string): void;
        addReferencedObjectPath(objectPath: ObjectPath): void;
        addReferencedObjectPaths(objectPaths: Array<ObjectPath>): void;
        addActionResultHandler(action: Action, resultHandler: IResultHandler): void;
        buildRequestMessageBody(): RequestMessageBody;
        processResponse(msg: ResponseMessageBody): void;
        invalidatePendingInvalidObjectPaths(): void;
    }
}
declare module OfficeExtension {
    interface LoadOption {
        select?: string;
        expand?: string;
        top?: number;
        skip?: number;
    }
    class ClientRequestContext {
        private m_nextId;
        private m_pendingRequest;
        private m_url;
        private m_references;
        _requestExecutor: IRequestExecutor;
        _rootObject: ClientObject;
        _customData: string;
        _processingResult: boolean;
        constructor(url?: string);
        _pendingRequest: ClientRequest;
        references: References;
        load(clientObj: ClientObject, option?: string | LoadOption): void;
        trace(message: string): void;
        private parseSelectExpand(select);
        private executeAsyncPrivate(doneCallback, failCallback);
        executeAsync(): IPromise<ClientRequestResult>;
        _nextId(): number;
    }
}
declare module OfficeExtension {
    enum ClientRequestFlags {
        None = 0,
        WriteOperation = 1,
    }
}
declare module OfficeExtension {
    class ClientRequestResult {
        errorCode: string;
        errorMessage: string;
        traceMessages: Array<string>;
        constructor(traceMessages: Array<string>);
    }
}
declare module OfficeExtension {
    class ClientResult<T> implements IResultHandler {
        private m_value;
        value: T;
        _handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    class Constants {
        static getItemAt: string;
        static id: string;
        static idPrivate: string;
        static index: string;
        static items: string;
        static iterativeExecutor: string;
        static localDocument: string;
        static localDocumentApiPrefix: string;
        static referenceId: string;
    }
}
declare module OfficeExtension {
    class Error {
        name: string;
        message: string;
        stack: string;
        constructor(message?: string);
    }
    class RuntimeError extends Error {
        traceMessages: Array<string>;
        debugInfo: {
            errorLocation?: string;
        };
        constructor(name: string, message: string, traceMessages: Array<string>, debugInfo: {
            errorLocation?: string;
        });
        errorCode: string;
        errorMessage: string;
        toString(): string;
    }
}
declare module OfficeExtension {
    class ErrorCodes {
        static accessDenied: string;
        static generalException: string;
    }
}
declare module OfficeExtension {
    class InstantiateActionResultHandler implements IResultHandler {
        private m_clientObject;
        constructor(clientObject: ClientObject);
        _handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    interface IPromise<R> {
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => void): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => void): IPromise<U>;
        catch<U>(onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        catch<U>(onRejected?: (error: any) => U): IPromise<U>;
        catch<U>(onRejected?: (error: any) => void): IPromise<U>;
    }
}
declare module OfficeExtension {
    interface IRequestExecutorRequestMessage {
        Url: string;
        Headers: {
            [name: string]: string;
        };
        Body: RequestMessageBody;
    }
    interface IRequestExecutorResponseMessage {
        ErrorCode: string;
        ErrorMessage: string;
        Headers: {
            [name: string]: string;
        };
        Body: ResponseMessageBody;
    }
    interface IRequestExecutor {
        executeAsync(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage, callback: (responseMessage: IRequestExecutorResponseMessage) => void): any;
    }
}
declare module OfficeExtension {
    interface IResultHandler {
        _handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    enum RichApiRequestMessageIndex {
        CustomData = 0,
        Method = 1,
        PathAndQuery = 2,
        Headers = 3,
        Body = 4,
        AppPermission = 5,
        RequestFlags = 6,
    }
    enum RichApiResponseMessageIndex {
        StatusCode = 0,
        Headers = 1,
        Body = 2,
    }
    enum ActionType {
        Instantiate = 1,
        Query = 2,
        Method = 3,
        SetProperty = 4,
        Trace = 5,
    }
    enum ObjectPathType {
        GlobalObject = 1,
        NewObject = 2,
        Method = 3,
        Property = 4,
        Indexer = 5,
        ReferenceId = 6,
    }
    interface ArgumentInfo {
        Arguments?: Array<any>;
        ReferencedObjectPathIds?: Array<number>;
    }
    interface QueryInfo {
        Select?: Array<string>;
        Expand?: Array<string>;
        Skip?: number;
        Top?: number;
    }
    interface ActionInfo {
        Id: number;
        ActionType: ActionType;
        Name: string;
        ObjectPathId: number;
        ArgumentInfo?: ArgumentInfo;
        QueryInfo?: QueryInfo;
    }
    interface ActionResult {
        ActionId: number;
        Value: any;
    }
    interface ObjectPathInfo {
        Id: number;
        ObjectPathType: ObjectPathType;
        Name: string;
        ParentObjectPathId?: number;
        ArgumentInfo?: ArgumentInfo;
    }
    interface RequestMessageBody {
        Actions: Array<ActionInfo>;
        ObjectPaths: {
            [Id: number]: ObjectPathInfo;
        };
    }
    interface ErrorInfo {
        Code: string;
        Message: string;
        Location: string;
    }
    interface ResponseMessageBody {
        Error: ErrorInfo;
        Results: Array<ActionResult>;
        TraceIds: Array<number>;
    }
}
declare module OfficeExtension {
    class ObjectPath {
        private m_objectPathInfo;
        private m_isWriteOperation;
        private m_parentObjectPath;
        private m_argumentObjectPaths;
        private m_isCollection;
        private m_isInvalidAfterRequest;
        private m_isValid;
        constructor(objectPathInfo: ObjectPathInfo, parentObjectPath: ObjectPath, isCollection: boolean, isInvalidAfterRequest: boolean);
        objectPathInfo: ObjectPathInfo;
        isWriteOperation: boolean;
        isCollection: boolean;
        isInvalidAfterRequest: boolean;
        parentObjectPath: ObjectPath;
        argumentObjectPaths: Array<ObjectPath>;
        isValid: boolean;
        updateUsingObjectData(value: Object): void;
    }
}
declare module OfficeExtension {
    class ObjectPathFactory {
        static createGlobalObjectObjectPath(context: ClientRequestContext): ObjectPath;
        static createNewObjectObjectPath(context: ClientRequestContext, typeName: string, isCollection?: boolean): ObjectPath;
        static createPropertyObjectPath(context: ClientRequestContext, parent: ClientObject, propertyName: string, isCollection?: boolean, isInvalidAfterRequest?: boolean): ObjectPath;
        static createIndexerObjectPath(context: ClientRequestContext, parent: ClientObject, args: Array<any>): ObjectPath;
        static createIndexerObjectPathUsingParentPath(context: ClientRequestContext, parentObjectPath: ObjectPath, args: Array<any>): ObjectPath;
        static createMethodObjectPath(context: ClientRequestContext, parent: ClientObject, methodName: string, operationType: OperationType, args: Array<any>, isCollection?: boolean, isInvalidAfterRequest?: boolean): ObjectPath;
        static createChildItemObjectPathUsingIndexer(context: ClientRequestContext, parent: ClientObject, childItem: Object): ObjectPath;
        static createChildItemObjectPathUsingGetItemAt(context: ClientRequestContext, parent: ClientObject, childItem: Object, index: number): ObjectPath;
        static createReferenceIdObjectPath(context: ClientRequestContext, referenceId: string): ObjectPath;
    }
}
declare module OfficeExtension {
    class OfficeJsRequestExecutor implements IRequestExecutor {
        private static OfficeJsErrorCode_ooeNoCapability;
        executeAsync(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage, callback: (responseMessage: IRequestExecutorResponseMessage) => void): void;
    }
}
declare module OfficeExtension {
    class OfficeXHRSettings {
        oldxhr: () => IXMLHttpRequest;
        executeRichApiRequestAsync: (message: Array<any>, callback: (result: OSF.DDA.RichApi.ExecuteRichApiRequestResult) => void) => void;
    }
    function resetXHRFactory(oldFactory: () => IXMLHttpRequest): () => IXMLHttpRequest;
    function officeXHRFactory(): OfficeXHR;
    class OfficeXHR implements IXMLHttpRequest {
        private static UNSENT;
        private static OPENED;
        private static DONE;
        static settings: OfficeXHRSettings;
        private m_innerXhr;
        private m_method;
        private m_url;
        private m_allResponseHeaders;
        private m_responseHeaders;
        private m_requestHeaders;
        onreadystatechange: () => void;
        readyState: number;
        status: number;
        statusText: string;
        response: any;
        responseText: string;
        responseType: string;
        open(method: string, url: string): void;
        abort(): void;
        send(body: string): void;
        setRequestHeader(header: string, value: string): void;
        getResponseHeader(header: string): string;
        getAllResponseHeaders(): string;
        overrideMimeType(mimeType: string): void;
        private innerXhrOnreadystatechage();
        private officeContextRequestCallback(result);
        private setAllResponseHeaders(allResponseHeaders);
    }
}
declare module OfficeExtension {
    function _EnsurePromise(): void;
}
declare module OfficeExtension {
    enum OperationType {
        Default = 0,
        Read = 1,
    }
}
declare module OfficeExtension {
    class References {
        private m_context;
        constructor(context: ClientRequestContext);
        add(clientObject: ClientObject): void;
        remove(clientObject: ClientObject): void;
        removeAll(): void;
    }
}
declare module OfficeExtension {
    class ResourceStrings {
        static invalidObjectPath: string;
        static propertyNotLoaded: string;
        static invalidRequestContext: string;
    }
}
declare module OfficeExtension {
    class RichApiMessageUtility {
        static buildRequestMessageSafeArray(customData: string, requestFlags: number, method: string, path: string, headers: {
            [name: string]: string;
        }, body: string): Array<any>;
        static getResponseBody(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): string;
        static getResponseHeaders(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): {
            [name: string]: string;
        };
        static getResponseBodyFromSafeArray(data: Array<any>): string;
        static getResponseHeadersFromSafeArray(data: Array<any>): {
            [name: string]: string;
        };
        static getResponseStatusCode(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): number;
        static getResponseStatusCodeFromSafeArray(data: Array<any>): number;
    }
}
declare module OfficeExtension {
    class Utility {
        static checkArgumentNull(value: any, name: string): void;
        static isNullOrUndefined(value: any): boolean;
        static isUndefined(value: any): boolean;
        static isNullOrEmptyString(value: string): boolean;
        static trim(str: string): string;
        static caseInsensitiveCompareString(str1: string, str2: string): boolean;
        static isReadonlyRestRequest(method: string): boolean;
        static setMethodArguments(context: ClientRequestContext, argumentInfo: ArgumentInfo, args: Array<any>): Array<ObjectPath>;
        static fixObjectPathIfNecessary(clientObject: ClientObject, value: Object): void;
        static validateObjectPath(clientObject: ClientObject): void;
        static validateReferencedObjectPaths(objectPaths: Array<ObjectPath>): void;
        static validateContext(context: ClientRequestContext, obj: ClientObject): void;
        static log(message: string): void;
        static load(clientObj: ClientObject, option?: string | LoadOption): void;
        private static s_underscoreCharCode;
        static throwError(resourceId: string, arg?: string): void;
        static throwIfNotLoaded(propertyName: string, fieldValue: any): void;
        static getObjectPathExpression(objectPath: ObjectPath): string;
        private static normalizeName(name);
    }
}
