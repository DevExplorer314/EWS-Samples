$("#run").click(function() {
  var contactName = "Catriona Abbott";
  var request = FindContactRequest(contactName);
  Office.context.mailbox.makeEwsRequestAsync(request, handleFindItemEwsResponse);
});

// This function builds the EWS request to find a contact by name.
function FindContactRequest(contactName) {
  const request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    '  <soap:Header><t:RequestServerVersion Version="Exchange2016" /></soap:Header>' +
    "  <soap:Body>" +
    '    <m:FindItem Traversal="Shallow">' +
    "      <m:ItemShape>" +
    "        <t:BaseShape>Default</t:BaseShape>" +
    "      </m:ItemShape >" +
    "      <m:Restriction>" +
    "        <t:IsEqualTo>" +
    '          <t:FieldURI FieldURI="contacts:DisplayName" />' +
    "          <t:FieldURIOrConstant>" +
    '            <t:Constant Value="' +
    contactName +
    '" />' +
    "          </t:FieldURIOrConstant>" +
    "        </t:IsEqualTo>" +
    "      </m:Restriction>" +
    "      <m:ParentFolderIds>" +
    '       <t:DistinguishedFolderId Id="contacts" />' +
    "      </m:ParentFolderIds>" +
    "    </m:FindItem>" +
    "  </soap:Body>" +
    "</soap:Envelope>";
  console.log(request);
  return request;
}

// This function handles the response from the EWS request by logging it to the console.
function handleFindItemEwsResponse(asyncResult) {
  if (asyncResult.status == "failed") {
    console.log(`Action failed with message: ${asyncResult.error.message}`);
    return;
  }

  // We got a response, so now we need to read the Id from that and send an UpdateItem request
  var response = asyncResult.value;
  console.log("EWS request succeeded. Response: " + response);
  var itemId = parseResponseForId(response);
  console.log("Contact item Id: " + itemId);
  var request = UpdateContactRequest(itemId, "test@test.com");
  Office.context.mailbox.makeEwsRequestAsync(request, handleUpdateItemEwsResponse);  
}

// This function handles the response from the EWS request by logging it to the console.
function handleUpdateItemEwsResponse(asyncResult) {
  if (asyncResult.status == "failed") {
    console.log(`Action failed with message: ${asyncResult.error.message}`);
    return;
  }

  // We got a response, just log that (hopefully success)
  var response = asyncResult.value;
  console.log("EWS request succeeded. Response: " + response);
}

// This function parses the response from the EWS request and returns the EWS Id of the contact.
function parseResponseForId(response) {
  // TODO: Implement this function to parse the response from the
  // FindContactRequest function and extract the EWS Id of the contact.
  // The exact implementation will depend on the format of the response.
  // For example:
  var itemId = $(response).find("t\\:ItemId").attr("Id");
  var changeKey = $(response).find("t\\:ItemId").attr("ChangeKey");
  var ewsId = '          <t:ItemId Id="' + itemId + '" ChangeKey="' + changeKey + '"/>';
  return ewsId;
}

// This function builds the EWS request to update the contact with the specified EWS Id.
function UpdateContactRequest(ewsId, newEmailAddress) {
  const request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    '  <soap:Header><t:RequestServerVersion Version="Exchange2016" /></soap:Header>' +
    "  <soap:Body>" +
    '    <m:UpdateItem ConflictResolution="AlwaysOverwrite" MessageDisposition="SaveOnly">' +
    "      <m:ItemChanges>" +
    "        <t:ItemChange>" +
    '          ' + ewsId +
    "          <t:Updates>" +
    "            <t:SetItemField>" +
    '              <t:IndexedFieldURI FieldURI="contacts:EmailAddress" FieldIndex="EmailAddress3" />' +
    "              <t:Contact>" +
    "                <t:EmailAddresses>" +
    '                  <t:Entry Key="EmailAddress3" >' + newEmailAddress + '</t:Entry>' +
    "                </t:EmailAddresses>" +
    "              </t:Contact>" +
    "            </t:SetItemField>" +
    "          </t:Updates>" +
    "        </t:ItemChange>" +
    "      </m:ItemChanges>" +
    "    </m:UpdateItem>" +
    "  </soap:Body>" +
    "</soap:Envelope>";

  console.log(request);
  return request;
}

// This function builds the EWS request to find a contact by name.
function updateContact(ewsId, newEmailAddress) {
  const request = UpdateContactRequest(ewsId, newEmailAddress);
  Office.context.mailbox.makeEwsRequestAsync(request, handleUpdateContactResponse);
}

// This function handles the response from the EWS request by logging it to the console.
function handleUpdateContactResponse(asyncResult) {
  if (asyncResult.status == "failed") {
    console.log(`Action failed with message: ${asyncResult.error.message}`);
  } else {
    var response = asyncResult.value;
    console.log("EWS request to update contact succeeded. Response: " + response);
  }
}

// This function parses the response from the EWS request and returns the EWS Id of the contact.
function handleEwsResponse(asyncResult) {
  if (asyncResult.status == "failed") {
    console.log(`Action failed with message: ${asyncResult.error.message}`);
  } else {
    var response = asyncResult.value;
    console.log("EWS request to find contact succeeded. Response: " + response);
    var ewsId = parseResponseForId(response);
    if (ewsId) {
      var newEmailAddress = "newemail@example.com"; // TODO: Replace this with the new email address.
      updateContact(ewsId, newEmailAddress);
    }
  }
}
