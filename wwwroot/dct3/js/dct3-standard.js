(function (window, $) {
  'use strict';

  // Namespace
  window.DCT3STD = window.DCT3STD || {};

  /**
   * Initialize a DataTable with sensible defaults.
   * @param {string|HTMLElement|jQuery} selector
   * @param {object} opts - DataTables options to merge
   * @returns {DataTable|undefined}
   */
  DCT3STD.initDataTable = function (selector, opts) {
    if (!$.fn || !$.fn.DataTable) {
      console.warn('DataTables plugin not loaded');
      return;
    }
    var defaults = {
      paging: true,
      lengthChange: true,
      pageLength: 10,
      searching: false,
      ordering: true,
      info: true,
      autoWidth: false
    };
    return $(selector).DataTable($.extend(true, defaults, opts || {}));
  };

  /**
   * Format date to dd/MM/yyyy if possible.
   */
  DCT3STD.formatDate = function (val) {
    if (!val) return '';
    var d = new Date(val);
    if (isNaN(d)) return val;
    var dd = ('0' + d.getDate()).slice(-2);
    var mm = ('0' + (d.getMonth() + 1)).slice(-2);
    var yy = d.getFullYear();
    return dd + '/' + mm + '/' + yy;
  };

})(window, window.jQuery);