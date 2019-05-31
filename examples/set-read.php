<?php

// Set connection information.
$host = '';
$username = '';
$password = '';
$version = Client::VERSION_2016;
$client = new Client($host, $username, $password, $version);

$request = new FindItemType();
$request->Traversal = ItemQueryTraversalType::SHALLOW; // This is indeed needed for Exchange 2007
$request->ParentFolderIds = new NonEmptyArrayOfBaseFolderIdsType();

// Return all message properties.
$request->ItemShape = new ItemResponseShapeType();
$request->ItemShape->BaseShape = DefaultShapeNamesType::ALL_PROPERTIES;
// Search in the user's inbox.
$folder_id = new DistinguishedFolderIdType();
$folder_id->Id = DistinguishedFolderIdNameType::INBOX;
$request->ParentFolderIds->DistinguishedFolderId[] = $folder_id;
$response = $client->FindItem($request);

// Iterate over the results, printing any error messages or message subjects.
$response_messages = $response->ResponseMessages->FindItemResponseMessage;
foreach ($response_messages as $response_message) {
    // Make sure the request succeeded.
    if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
        $code = $response_message->ResponseCode;
        $message = $response_message->MessageText;
        fwrite(
            STDERR,
            "Failed to search for messages with \"$code: $message\"\n"
        );
        continue;
    }
    // Iterate over the messages that were found, printing the subject for each.
    $items = $response_message->RootFolder->Items->Message;
    foreach ($items as $item) {
    	//var_dump($item);
        if($item->IsRead) {
          continue;
        }
        $subject = $item->Subject;
        $id = $item->ItemId->Id;
        $item->IsRead = true;
        
        // create a new request type
        $update = new UpdateItemType();
        $update->ConflictResolution = 'AlwaysOverwrite';
        $update->MessageDisposition = MessageDispositionType::SAVE_ONLY;
        $update->ItemChanges = new NonEmptyArrayOfItemChangesType();

        $itemChange = new ItemChangeType();
        $itemChange->ItemId = $item->ItemId;

        $itemChange->Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

        $setItemField = new SetItemFieldType();
        $setItemField->FieldURI = new PathToUnindexedFieldType();
        $setItemField->FieldURI->FieldURI = 'message:IsRead';
        $setItemField->Message = new MessageType();
        $setItemField->Message->IsRead = true;
        $itemChange->Updates->SetItemField = $setItemField;
        $update->ItemChanges->ItemChange[] = $itemChange;
        $updateResponse = $client->UpdateItem($update);

        fwrite(STDOUT, "$subject (now read): $id\n");
    }
}
