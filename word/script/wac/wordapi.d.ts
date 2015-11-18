declare module Word.Utility {
    function _normalizeSearchOptions(context: OfficeExtension.ClientRequestContext, searchOptions: any): Word.SearchOptions;
}
declare module Word {
    /**
     *
     * Represents the body of a document or a section.
     */
    class Body extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_inlinePictures;
        private m_paragraphs;
        private m_parentContentControl;
        private m_style;
        private m_text;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of rich text content control objects that are in the body. Read-only.
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the body. Use this to get and set font name, size, color, and other properties. Read-only.
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects that are in the body. The collection does not include floating images. Read-only.
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of paragraph objects that are in the body. Read-only.
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the content control that contains the body. Returns null if there isn't a parent content control. Read-only.
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets or sets the style used for the body. This is the name of the pre-installed or custom style.
         */
        style: string;
        /**
         *
         * Gets the text of the body. Use the insertText method to insert text. Read-only.
         */
        text: string;
        _ReferenceId: string;
        /**
         *
         * Clears the contents of the body object. The user can perform the undo operation on the cleared content.
         *
         */
        clear(): void;
        /**
         *
         * Gets the HTML representation of the body object.
         *
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the OOXML (Office Open XML) representation of the body object.
         *
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.
         *
         * @param breakType Required. The break type to add to the body.
         * @param insertLocation Required. The value can be 'Start' or 'End'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the body object with a Rich Text content control.
         *
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param html Required. The HTML to be inserted in the document.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted in the body.
         * @param insertLocation Required. The value can be 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param ooxml Required. The OOXML to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Start' or 'End'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param text Required. Text to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSoundsLike?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildCards?: boolean;
        }): Word.SearchResultCollection;
        /**
         *
         * Selects the body and navigates the Word UI to it.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Body;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     */
    class ContentControl extends OfficeExtension.ClientObject {
        private m_appearance;
        private m_cannotDelete;
        private m_cannotEdit;
        private m_color;
        private m_contentControls;
        private m_font;
        private m_id;
        private m_inlinePictures;
        private m_paragraphs;
        private m_parentContentControl;
        private m_placeholderText;
        private m_removeWhenEdited;
        private m_style;
        private m_tag;
        private m_text;
        private m_title;
        private m_type;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of content control objects in the content control. Read-only.
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Get the collection of paragraph objects in the content control. Read-only.
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the content control that contains the content control. Returns null if there isn't a parent content control. Read-only.
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'.
         */
        appearance: string;
        /**
         *
         * Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
         */
        cannotDelete: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the user can edit the contents of the content control.
         */
        cannotEdit: boolean;
        /**
         *
         * Gets or sets the color of the content control. Color is set in "#RRGGBB" format or by using the color name.
         */
        color: string;
        /**
         *
         * Gets an integer that represents the content control identifier. Read-only.
         */
        id: number;
        /**
         *
         * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
         */
        placeholderText: string;
        /**
         *
         * Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
         */
        removeWhenEdited: boolean;
        /**
         *
         * Gets or sets the style used for the content control. This is the name of the pre-installed or custom style.
         */
        style: string;
        /**
         *
         * Gets or sets a tag to identify a content control.
         */
        tag: string;
        /**
         *
         * Gets the text of the content control. Read-only.
         */
        text: string;
        /**
         *
         * Gets or sets the title for a content control.
         */
        title: string;
        /**
         *
         * Gets the content control type. Only rich text content controls are supported currently. Read-only.
         */
        type: string;
        _ReferenceId: string;
        /**
         *
         * Clears the contents of the content control. The user can perform the undo operation on the cleared content.
         *
         */
        clear(): void;
        /**
         *
         * Deletes the content control and its content. If keepContent is set to true, the content is not deleted.
         *
         * @param keepContent Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.
         */
        delete(keepContent: boolean): void;
        /**
         *
         * Gets the HTML representation of the content control object.
         *
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the Office Open XML (OOXML) representation of the content control object.
         *
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
         *
         * @param breakType Required. Type of break.
         * @param insertLocation Required. The value can be 'Start', 'End', 'Before' or 'After'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param html Required. The HTML to be inserted in to the content control.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted in the content control.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param ooxml Required. The OOXML to be inserted in to the content control.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragrph text to be inserted.
         * @param insertLocation Required. The value can be 'Start', 'End', 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param text Required. The text to be inserted in to the content control.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSoundsLike?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildCards?: boolean;
        }): Word.SearchResultCollection;
        /**
         *
         * Selects the content control. This causes Word to scroll to the selection.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.ContentControl;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Contains a collection of ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     */
    class ContentControlCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.ContentControl>;
        _ReferenceId: string;
        /**
         *
         * Gets a content control by its identifier.
         *
         * @param id Required. A content control identifier.
         */
        getById(id: number): Word.ContentControl;
        /**
         *
         * Gets the content controls that have the specified tag.
         *
         * @param tag Required. A tag set on a content control.
         */
        getByTag(tag: string): Word.ContentControlCollection;
        /**
         *
         * Gets the content controls that have the specified title.
         *
         * @param title Required. The title of a content control.
         */
        getByTitle(title: string): Word.ContentControlCollection;
        /**
         *
         * Gets a content control by its index in the collection.
         *
         * @param index The index
         */
        getItem(index: number): Word.ContentControl;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.ContentControlCollection;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
     */
    class Document extends OfficeExtension.ClientObject {
        private m_body;
        private m_contentControls;
        private m_saved;
        private m_sections;
        /**
         *
         * Gets the body of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
         */
        body: Word.Body;
        /**
         *
         * Gets the collection of content control objects that are in the current document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the collection of section objects that are in the document. Read-only.
         */
        sections: Word.SectionCollection;
        /**
         *
         * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
         */
        saved: boolean;
        /**
         *
         * Gets the current selection of the document. Multiple selections are not supported.
         *
         */
        getSelection(): Word.Range;
        /**
         *
         * Saves the document. This will use the Word default file naming convention if the document has not been saved before.
         *
         */
        save(): void;
        _GetObjectByReferenceId(referenceId: string): OfficeExtension.ClientResult<any>;
        _GetObjectTypeNameByReferenceId(referenceId: string): OfficeExtension.ClientResult<string>;
        _RemoveAllReferences(): void;
        _RemoveReference(referenceId: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Document;
    }
    /**
     *
     * Represents a font.
     */
    class Font extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_doubleStrikeThrough;
        private m_highlightColor;
        private m_italic;
        private m_name;
        private m_size;
        private m_strikeThrough;
        private m_subscript;
        private m_superscript;
        private m_underline;
        private m__ReferenceId;
        /**
         *
         * Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
         */
        bold: boolean;
        /**
         *
         * Gets or sets the color for the specified font. You can provide the value in the "#RRGGBB" format or the color name.
         */
        color: string;
        /**
         *
         * Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.
         */
        doubleStrikeThrough: boolean;
        /**
         *
         * Gets or sets the highlight color for the specified font. You can provide the value as either in the "#RRGGBB" format or the color name.
         */
        highlightColor: string;
        /**
         *
         * Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
         */
        italic: boolean;
        /**
         *
         * Gets or sets a value that represents the name of the font.
         */
        name: string;
        /**
         *
         * Gets or sets a value that represents the font size in points.
         */
        size: number;
        /**
         *
         * Gets or sets a value that indicates whether the font has a strike through. True if the font is formatted as strikethrough text, otherwise, false.
         */
        strikeThrough: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
         */
        subscript: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
         */
        superscript: boolean;
        /**
         *
         * Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
         */
        underline: string;
        _ReferenceId: string;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Font;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Represents an inline picture.
     */
    class InlinePicture extends OfficeExtension.ClientObject {
        private m_altTextDescription;
        private m_altTextTitle;
        private m_height;
        private m_hyperlink;
        private m_lockAspectRatio;
        private m_paragraph;
        private m_parentContentControl;
        private m_width;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the paragraph that contains the inline image. Read-only.
         */
        paragraph: Word.Paragraph;
        /**
         *
         * Gets the content control that contains the inline image. Returns null if there isn't a parent content control. Read-only.
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets or sets a string that represents the alternative text associated with the inline image
         */
        altTextDescription: string;
        /**
         *
         * Gets or sets a string that contains the title for the inline image.
         */
        altTextTitle: string;
        /**
         *
         * Gets or sets a number that describes the height of the inline image.
         */
        height: number;
        /**
         *
         * Gets or sets the hyperlink associated with the inline image.
         */
        hyperlink: string;
        /**
         *
         * Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
         */
        lockAspectRatio: boolean;
        /**
         *
         * Gets or sets a number that describes the width of the inline image.
         */
        width: number;
        _Id: number;
        _ReferenceId: string;
        /**
         *
         * Deletes the inline picture from the document.
         *
         */
        delete(): void;
        /**
         *
         * Gets the base64 encoded string representation of the inline image.
         *
         */
        getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * @param breakType Required. The break type to add.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the inline picture with a rich text content control.
         *
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param html Required. The HTML to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Before' or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.
         *
         * @param ooxml Required. The OOXML to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param text Required. Text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.InlinePicture;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Contains a collection of [inlinePicture](inlinePicture.md) objects.
     */
    class InlinePictureCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.InlinePicture>;
        _ReferenceId: string;
        /**
         *
         * Gets an inline picture object by its index in the collection.
         *
         * @param index A number that identifies the index location of an inline picture object.
         */
        _GetItem(index: number): Word.InlinePicture;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.InlinePictureCollection;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Represents a single paragraph in a selection, range, content control, or document body.
     */
    class Paragraph extends OfficeExtension.ClientObject {
        private m_alignment;
        private m_contentControls;
        private m_firstLineIndent;
        private m_font;
        private m_inlinePictures;
        private m_leftIndent;
        private m_lineSpacing;
        private m_lineUnitAfter;
        private m_lineUnitBefore;
        private m_outlineLevel;
        private m_parentContentControl;
        private m_rightIndent;
        private m_spaceAfter;
        private m_spaceBefore;
        private m_style;
        private m_text;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of content control objects that are in the paragraph. Read-only.
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects that are in the paragraph. The collection does not include floating images. Read-only.
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the content control that contains the paragraph. Returns null if there isn't a parent content control. Read-only.
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets or sets the alignment for a paragraph. The value can  be "left", "centered", "right", or "justified".
         */
        alignment: string;
        /**
         *
         * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
         */
        firstLineIndent: number;
        /**
         *
         * Gets or sets the left indent value, in points, for the paragraph.
         */
        leftIndent: number;
        /**
         *
         * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
         */
        lineSpacing: number;
        /**
         *
         * Gets or sets the amount of spacing, in grid lines. after the paragraph.
         */
        lineUnitAfter: number;
        /**
         *
         * Gets or sets the amount of spacing, in grid lines, before the paragraph.
         */
        lineUnitBefore: number;
        /**
         *
         * Gets or sets the outline level for the paragraph.
         */
        outlineLevel: number;
        /**
         *
         * Gets or sets the right indent value, in points, for the paragraph.
         */
        rightIndent: number;
        /**
         *
         * Gets or sets the spacing, in points, after the paragraph.
         */
        spaceAfter: number;
        /**
         *
         * Gets or sets the spacing, in points, before the paragraph.
         */
        spaceBefore: number;
        /**
         *
         * Gets or sets the style used for the paragraph. This is the name of the pre-installed or custom style.
         */
        style: string;
        /**
         *
         * Gets the text of the paragraph. Read-only.
         */
        text: string;
        _Id: number;
        _ReferenceId: string;
        /**
         *
         * Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.
         *
         */
        clear(): void;
        /**
         *
         * Deletes the paragraph and its content from the document.
         *
         */
        delete(): void;
        /**
         *
         * Gets the HTML representation of the paragraph object.
         *
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the Office Open XML (OOXML) representation of the paragraph object.
         *
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * @param breakType Required. The break type to add to the document.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the paragraph object with a rich text content control.
         *
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param html Required. The HTML to be inserted in the paragraph.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param ooxml Required. The OOXML to be inserted in the paragraph.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param text Required. Text to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSoundsLike?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildCards?: boolean;
        }): Word.SearchResultCollection;
        /**
         *
         * Selects and navigates the Word UI to the paragraph.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Paragraph;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Contains a collection of [paragraph](paragraph.md) objects.
     */
    class ParagraphCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Paragraph>;
        _ReferenceId: string;
        /**
         *
         * Gets a paragraph object by its index in the collection.
         *
         * @param index A number that identifies the index location of a paragraph object.
         */
        _GetItem(index: number): Word.Paragraph;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.ParagraphCollection;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Represents a contiguous area in a document.
     */
    class Range extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_inlinePictures;
        private m_paragraphs;
        private m_parentContentControl;
        private m_style;
        private m_text;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of content control objects that are in the range. Read-only.
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inline picture objects that are in the range. Read-only.
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of paragraph objects that are in the range. Read-only.
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the content control that contains the range. Returns null if there isn't a parent content control. Read-only.
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets or sets the style used for the range. This is the name of the pre-installed or custom style.
         */
        style: string;
        /**
         *
         * Gets the text of the range. Read-only.
         */
        text: string;
        _Id: number;
        _ReferenceId: string;
        /**
         *
         * Clears the contents of the range object. The user can perform the undo operation on the cleared content.
         *
         */
        clear(): void;
        /**
         *
         * Deletes the range and its content from the document.
         *
         */
        delete(): void;
        /**
         *
         * Gets the HTML representation of the range object.
         *
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the OOXML representation of the range object.
         *
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Replace', 'Before' or 'After'.
         *
         * @param breakType Required. The break type to add.
         * @param insertLocation Required. The value can be 'Replace', 'Before' or 'After'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the range object with a rich text content control.
         *
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param html Required. The HTML to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param ooxml Required. The OOXML to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param text Required. Text to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSoundsLike?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildCards?: boolean;
        }): Word.SearchResultCollection;
        /**
         *
         * Selects and navigates the Word UI to the range.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Range;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Specifies the options to be included in a search operation.
     */
    class SearchOptions extends OfficeExtension.ClientObject {
        private m_ignorePunct;
        private m_ignoreSpace;
        private m_matchCase;
        private m_matchPrefix;
        private m_matchSoundsLike;
        private m_matchSuffix;
        private m_matchWholeWord;
        private m_matchWildCards;
        /**
         *
         * Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
         */
        ignorePunct: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to ignore all white space between words. Corresponds to the Ignore white-space characters check box in the Find and Replace dialog box.
         */
        ignoreSpace: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).
         */
        matchCase: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
         */
        matchPrefix: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to find words that sound similar to the search string. Corresponds to the Sounds like check box in the Find and Replace dialog box
         */
        matchSoundsLike: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
         */
        matchSuffix: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
         */
        matchWholeWord: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
         */
        matchWildCards: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.SearchOptions;
        /**
         * Create a new instance of Word.SearchOptions object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.SearchOptions;
    }
    /**
     *
     * Contains a collection of [range](range.md) objects as a result of a search operation.
     */
    class SearchResultCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Range>;
        _ReferenceId: string;
        /**
         *
         * Gets a range object by its index in the collection.
         *
         * @param index A number that identifies the index location of a range object.
         */
        _GetItem(index: number): Word.Range;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.SearchResultCollection;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Represents a section in a Word document.
     */
    class Section extends OfficeExtension.ClientObject {
        private m_body;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the body of the section. This does not include the header/footer and other section metadata. Read-only.
         */
        body: Word.Body;
        _Id: number;
        _ReferenceId: string;
        /**
         *
         * Gets one of the section's footers.
         *
         * @param type Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'.
         */
        getFooter(type: string): Word.Body;
        /**
         *
         * Gets one of the section's headers.
         *
         * @param type Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'.
         */
        getHeader(type: string): Word.Body;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Section;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * Contains the collection of the document's [section](section.md) objects.
     */
    class SectionCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Section>;
        _ReferenceId: string;
        /**
         *
         * Gets a section object by its index in the collection.
         *
         * @param index A number that identifies the index location of a section object.
         */
        _GetItem(index: number): Word.Section;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.SectionCollection;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * ContentControl types
     */
    module ContentControlType {
        var richText: string;
    }
    /**
     *
     * ContentControl appearance
     */
    module ContentControlAppearance {
        var boundingBox: string;
        var tags: string;
        var hidden: string;
    }
    /**
     *
     * Underline types
     */
    module UnderlineType {
        var none: string;
        var single: string;
        var word: string;
        var double: string;
        var dotted: string;
        var hidden: string;
        var thick: string;
        var dashLine: string;
        var dotLine: string;
        var dotDashLine: string;
        var twoDotDashLine: string;
        var wave: string;
    }
    module BreakType {
        var page: string;
        var column: string;
        var next: string;
        var sectionContinuous: string;
        var sectionEven: string;
        var sectionOdd: string;
        var line: string;
        var lineClearLeft: string;
        var lineClearRight: string;
        var textWrapping: string;
    }
    /**
     *
     * The insertion location types
     */
    module InsertLocation {
        var before: string;
        var after: string;
        var start: string;
        var end: string;
        var replace: string;
    }
    module Alignment {
        var unknown: string;
        var left: string;
        var centered: string;
        var right: string;
        var justified: string;
    }
    module HeaderFooterType {
        var primary: string;
        var firstPage: string;
        var evenPages: string;
    }
    module SelectionMode {
        var select: string;
        var start: string;
        var end: string;
    }
    module ErrorCodes {
        var accessDenied: string;
        var generalException: string;
        var invalidArgument: string;
        var itemNotFound: string;
        var notImplemented: string;
    }
}
declare module Word {
    class RequestContext extends OfficeExtension.ClientRequestContext {
        private m_document;
        constructor(url?: string);
        document: Document;
    }
}
