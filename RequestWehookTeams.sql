/*
Script to send a request to the Teams webhook (needs to be created in advance)

*/





DECLARE @authHeader NVARCHAR(64);
DECLARE @contentType NVARCHAR(64);
DECLARE @postData NVARCHAR(2000);
DECLARE @responseText NVARCHAR(2000);
DECLARE @responseXML NVARCHAR(2000);
DECLARE @ret INT;
DECLARE @status NVARCHAR(32);
DECLARE @statusText NVARCHAR(32);
DECLARE @token INT;
DECLARE @url NVARCHAR(256);

SET @contentType = N'application/json';
SET @postData = N'{
        "text": "Mensagem Teste SS"
    }';
    
    --Set your url from webhook here
SET @url = N'https://..........';

-- Open the connection.

-- Send the request.
EXEC @ret = [sys].[sp_OAMethod]
    NULL
   ,'open'
   ,NULL
   ,'POST'
   ,@url
   ,'false';

EXEC @ret = [sys].[sp_OAMethod]
    NULL
   ,'setRequestHeader'
   ,NULL
   ,'Authentication'
   ,@authHeader;

EXEC @ret = [sys].[sp_OAMethod]
    NULL
   ,'setRequestHeader'
   ,NULL
   ,'Content-type'
   ,@contentType;

EXEC @ret = [sys].[sp_OAMethod]
    NULL
   ,'send'
   ,NULL
   ,@postData;

-- Handle the response.
EXEC @ret = [sys].[sp_OAGetProperty]
    @token
   ,'status'
   ,@status OUT;

EXEC @ret = [sys].[sp_OAGetProperty]
    @token
   ,'statusText'
   ,@statusText OUT;

EXEC @ret = [sys].[sp_OAGetProperty]
    @token
   ,'responseText'
   ,@responseText OUT;

-- Show the response.
PRINT 'Status: ' + @status + ' (' + @statusText + ')';
PRINT 'Response text: ' + @responseText;

-- Close the connection.
EXEC @ret = [sys].[sp_OADestroy]
    @token;

IF @ret <> 0
    RAISERROR('Unable to close HTTP connection.', 10, 1);