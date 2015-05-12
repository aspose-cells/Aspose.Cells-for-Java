function showAboutDialog() {
    PF('aboutDialog').show();
    return false;
}

function showFileUploadDialog() {
    PF('fileUploadDialog').chooseButton.find('input[type=file]').click();
    return false;
}

function showColumnWidthDialog() {
    PF('columnWidthDialog').show();
    return false;
}

function hideColumnWidthDialog() {
    PF('columnWidthDialog').hide();
    return false;
}

function showRowHeightDialog() {
    PF('rowHeightDialog').show();
    return false;
}

function hideRowHeightDialog() {
    PF('rowHeightDialog').hide();
    return false;
}

function saveFormulaBarContents() {
    var newContents = PF('formulaBar').jq.val();
    $(sheet_datatable.selectedCell).find('.ui-cell-editor-input input').val(newContents);
    sheet_datatable.saveCell($(sheet_datatable.selectedCell));

    return false;
}

function showInsertFunctionDialog() {
    PF('insertFunctionDialog').show();
    return false;
}

function hideInsertFunctionDialog() {
    PF('insertFunctionDialog').hide();
    return false;
}

function showSupportedFunctionsPage() {
    PF('insertFunctionDialog').hide();
    window.open('http://www.aspose.com/docs/display/cellsjava/Supported+Formula+Functions');
    return false;
}

var sheet_datatable;
PrimeFaces.widget.DataTable = PrimeFaces.widget.DataTable.extend({
    bindEditEvents: function() {
        sheet_datatable = this;
        var $this = this;
        // From primefaces' datatable.js

        var cellSelector = '> tr > td.ui-editable-column';
        this.tbody.off('click.datatable-cell', cellSelector)
                .on('click.datatable-cell', cellSelector, null, function(e) {
                    singleCellSelectionHandler($this, this);
                });

        this.tbody.off('dblclick.datatable-cell', cellSelector)
                .on('dblclick.datatable-cell', cellSelector, null, function(e) {
                    $this.incellClick = true;
                    var cell = $(this);
                    if (!cell.hasClass('ui-cell-editing')) {
                        $this.showCellEditor($(this));
                        var inp = $(this).find('.ui-cell-editor-input input')[0];
                        inp.setSelectionRange(inp.value.length + 1, inp.value.length + 1);
                    }
                });

        $(document).off('click.datatable-cell-blur' + this.id)
                .on('click.datatable-cell-blur' + this.id, function(e) {
                    if (!$this.incellClick && $this.currentCell && !$this.contextMenuClick) {
                        $this.saveCell($this.currentCell);
                    }
                    $this.incellClick = false;
                    $this.contextMenuClick = false;
                });

    }
});

function updateCurrentCells() {
    var columnId = PF('currentColumnIdValue').jq.val();
    var rowId = PF('currentRowIdValue').jq.val();
    var sel = '.ui-datatable .ui-cell-editor-input input[data-columnid=' + columnId + '][data-rowid=' + rowId + ']';

    updatePartialView([
        {name: 'id', value: $(sel).closest('.ui-cell-editor').attr('id')}
    ]);
}

function singleCellSelectionHandler(datatable, cell) {
    var columnId = $(cell).find('.ui-cell-editor-input input').attr('data-columnid');
    var rowId = $(cell).find('.ui-cell-editor-input input').attr('data-rowid');
    var cellName = $(cell).find('.ui-cell-editor-input input').attr('data-cellname');
    var clientId = $(cell).find('.ui-cell-editor').attr('id');

    PF('currentColumnIdValue').jq.val(columnId);
    PF('currentRowIdValue').jq.val(rowId);
    PF('currentCellNameValue').jq.val(cellName);
    PF('currentCellClientIdValue').jq.val(clientId);
    if ($(cell).find('.ui-cell-editor-output div').hasClass('b')) {
        PF('boldOptionButton').check();
    } else {
        PF('boldOptionButton').uncheck();
    }
    if ($(cell).find('.ui-cell-editor-output div').hasClass('i')) {
        PF('italicOptionButton').check();
    } else {
        PF('italicOptionButton').uncheck();
    }
    if ($(cell).find('.ui-cell-editor-output div').hasClass('u')) {
        PF('underlineOptionButton').check();
    } else {
        PF('underlineOptionButton').uncheck();
    }
    var cellFont = $(cell).find('.ui-cell-editor-output div').css('font-family').slice(1, -1);
    if (cellFont) {
        PF('fontOptionSelector').selectValue(cellFont);
    }
    var cellFontSize = $(cell).find('.ui-cell-editor-output div')[0].style.fontSize.replace('pt', '');
    if (cellFontSize) {
        PF('fontSizeOptionSelector').selectValue(cellFontSize);
    }
    ['al', 'ac', 'ar', 'aj'].forEach(function(v) {
        if ($(cell).find('.ui-cell-editor-output div').hasClass(v)) {
            // TODO: save the value to PF('alignOptionSelector');
        }
    });
    PF('currentColumnWidthValue').jq.val($(cell).width());
    PF('currentRowHeightValue').jq.val($(cell).height());
    PF('formulaBar').jq.val($(cell).find('.ui-cell-editor-input input').val());

    $(datatable.selectedCell).removeClass('sheet-selected-cell');
    $(cell).addClass('sheet-selected-cell');
    datatable.selectedCell = cell;
}

setInterval(function() {

    try {
        $('#sheet .ui-datatable').height($(window).height() - $('#sheet .ui-datatable').position().top - 1);
    } catch (x) {
    }

    try {
        var _osidufasiuf = PF('fontColorSelector').cfg.onHide;
        PF('fontColorSelector').overlay.data('colorpicker').onHide = function(e) {
            _osidufasiuf.apply(this, [arguments]);
            PF('applyFormattingButton').getJQ().click();
        };
    } catch (x) {
    }

    try {
        var _njksiuhgvbd = PF('fillColorSelector').cfg.onHide;
        PF('fillColorSelector').overlay.data('colorpicker').onHide = function(e) {
            _njksiuhgvbd.apply(this, [arguments]);
            PF('applyFormattingButton').getJQ().click();
        };
    } catch (x) {
    }

}, 786);

$(document).ready(function() {
    [
        '.ui-button.cell-formatting',
        '.ui-selectonemenu-panel.cell-formatting .ui-selectonemenu-list-item',
        '.ui-selectonebutton.cell-formatting .ui-button',
    ].forEach(function(sel) {
        $(document).on('click', sel, function(e) {
            PF('applyFormattingButton').getJQ().click();
        });
    });
});

