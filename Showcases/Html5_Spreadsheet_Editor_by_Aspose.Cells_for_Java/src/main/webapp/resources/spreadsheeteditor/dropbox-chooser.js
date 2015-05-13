function showDropboxChooser() {
    if (!Dropbox.isBrowserSupported()) {
        alert("Your browser is not supported to use this feature.");
        return false;
    }

    Dropbox.choose({
        success: function(files) {
            var link = files[0].link;
            PF('dropboxChooserSelectedUrl').jq.val(link);
            PF('dropboxChooserSelectedUrlApply').jq.trigger('click');
        },
        linkType: "direct",
        extensions: ['.xlsx', '.xls', '.ods'],
        multiselect: false
    });

    return false;
}

