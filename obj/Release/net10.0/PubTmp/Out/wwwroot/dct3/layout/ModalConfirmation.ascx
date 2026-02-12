<%@ Control Language="VB" AutoEventWireup="false" %>
<!-- Confirmation Modal Template -->
<link rel="stylesheet" href="DCT3_Standarisasi/Layout/notifications.css" />
<div id="dct-confirmation-overlay" class="dct-overlay" aria-hidden="true">
  <div class="dct-modal" role="dialog" aria-labelledby="dct-confirmation-title" aria-modal="true">
    <button type="button" class="dct-close" title="Close">
      <img src="DCT3_Standarisasi/assets/icons/CloseNotification.png" alt="Close" />
    </button>
    <div class="dct-header">
      <img src="DCT3_Standarisasi/assets/icons/Confirmation.png" alt="Confirmation" />
      <div id="dct-confirmation-title" class="dct-title">Confirmation</div>
    </div>
    <div id="dct-confirmation-message" class="dct-body"></div>
    <div class="dct-actions">
      <button type="button" class="dct-btn dct-btn-primary" data-action="yes">Yes</button>
      <button type="button" class="dct-btn" data-action="no">No</button>
    </div>
  </div>
  <script src="DCT3_Standarisasi/Layout/notifications.js"></script>
  <!-- Usage: DCT3Modals.open('confirmation', { message: 'text', onYes: fn, onNo: fn }); -->
</div>