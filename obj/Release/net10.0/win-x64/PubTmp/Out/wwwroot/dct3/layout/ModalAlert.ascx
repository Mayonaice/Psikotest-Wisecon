<%@ Control Language="VB" AutoEventWireup="false" %>
<!-- Alert Modal Template -->
<link rel="stylesheet" href="DCT3_Standarisasi/Layout/notifications.css" />
<div id="dct-alert-overlay" class="dct-overlay" aria-hidden="true">
  <div class="dct-modal" role="dialog" aria-labelledby="dct-alert-title" aria-modal="true">
    <button type="button" class="dct-close" title="Close">
      <img src="DCT3_Standarisasi/assets/icons/CloseNotification.png" alt="Close" />
    </button>
    <div class="dct-header">
      <img src="DCT3_Standarisasi/assets/icons/Alert.png" alt="Alert" />
      <div id="dct-alert-title" class="dct-title">Alert</div>
    </div>
    <div id="dct-alert-message" class="dct-body"></div>
    <div class="dct-actions">
      <button type="button" class="dct-btn dct-btn-primary" data-action="ok">OK</button>
    </div>
  </div>
  <script src="DCT3_Standarisasi/Layout/notifications.js"></script>
  <!-- Usage: DCT3Modals.open('alert', { message: 'text', html:true, onOk: fn }); -->
</div>