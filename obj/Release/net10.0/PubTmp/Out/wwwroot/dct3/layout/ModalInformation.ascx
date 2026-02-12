<%@ Control Language="VB" AutoEventWireup="false" %>
<!-- Information Modal Template -->
<link rel="stylesheet" href="DCT3_Standarisasi/Layout/notifications.css" />
<div id="dct-information-overlay" class="dct-overlay" aria-hidden="true">
  <div class="dct-modal" role="dialog" aria-labelledby="dct-information-title" aria-modal="true">
    <button type="button" class="dct-close" title="Close">
      <img src="DCT3_Standarisasi/assets/icons/CloseNotification.png" alt="Close" />
    </button>
    <div class="dct-header">
      <img src="DCT3_Standarisasi/assets/icons/Information.png" alt="Information" />
      <div id="dct-information-title" class="dct-title">Information</div>
    </div>
    <div id="dct-information-message" class="dct-body"></div>
    <div class="dct-actions">
      <button type="button" class="dct-btn dct-btn-primary" data-action="ok">OK</button>
    </div>
  </div>
  <script src="DCT3_Standarisasi/Layout/notifications.js"></script>
  <!-- Usage: DCT3Modals.open('information', { message: 'text', onOk: fn }); -->
</div>