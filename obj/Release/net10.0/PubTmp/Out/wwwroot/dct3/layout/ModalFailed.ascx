<%@ Control Language="VB" AutoEventWireup="false" %>
<!-- Failed Modal Template -->
<link rel="stylesheet" href="DCT3_Standarisasi/Layout/notifications.css" />
<div id="dct-failed-overlay" class="dct-overlay" aria-hidden="true">
  <div class="dct-modal" role="dialog" aria-labelledby="dct-failed-title" aria-modal="true">
    <button type="button" class="dct-close" title="Close">
      <img src="DCT3_Standarisasi/assets/icons/CloseNotification.png" alt="Close" />
    </button>
    <div class="dct-header">
      <img src="DCT3_Standarisasi/assets/icons/Failed.png" alt="Failed" />
      <div id="dct-failed-title" class="dct-title">Failed</div>
    </div>
    <div id="dct-failed-message" class="dct-body"></div>
    <div class="dct-actions">
      <button type="button" class="dct-btn dct-btn-danger" data-action="ok">OK</button>
    </div>
  </div>
  <script src="DCT3_Standarisasi/Layout/notifications.js"></script>
  <!-- Usage: DCT3Modals.open('failed', { message: 'text', html:true, onOk: fn }); -->
</div>