Attribute VB_Name = "MErrors"
Option Explicit

Public Const ERROR_SSH_INVALID_IDENTIFICATION_STRING As Long = &H1          'Invalid identification string of SSH-protocol.
Public Const ERROR_SSH_INVALID_VERSION As Long = &H2                        'Invalid or unsupported version.
Public Const ERROR_SSH_INVALID_MESSAGE_CODE As Long = &H3                   'Unsupported message code.
Public Const ERROR_SSH_INVALID_CRC As Long = &H4                            'Message CRC is invalid.
Public Const ERROR_SSH_INVALID_PACKET_TYPE As Long = &H5                    'Invalid (unknown) packet type.
Public Const ERROR_SSH_INVALID_PACKET As Long = &H6                         'Packet composed incorrectly.
Public Const ERROR_SSH_UNSUPPORTED_CIPHER As Long = &H7                     'There is no cipher supported by both: client and server.
Public Const ERROR_SSH_UNSUPPORTED_AUTH_TYPE As Long = &H8                  'Authentication type is unsupported.
Public Const ERROR_SSH_INVALID_RSA_CHALLENGE As Long = &H9                  'The wrong signature during public key-authentication.
Public Const ERROR_SSH_AUTHENTICATION_FAILED As Long = &HA                  'Authentication failed. There could be wrong password or something else.
Public Const ERROR_SSH_INVALID_PACKET_SIZE As Long = &HB                    'The packet is too large.
Public Const ERROR_SSH_HOST_NOT_ALLOWED_TO_CONNECT As Long = &H65           'Connection was rejected by remote host.
Public Const ERROR_SSH_PROTOCOL_ERROR As Long = &H66                        'Another protocol error.
Public Const ERROR_SSH_KEY_EXCHANGE_FAILED As Long = &H67                   'Key exchange failed.
Public Const ERROR_SSH_INVALID_MAC As Long = &H69                           'Received packet has invalid MAC.
Public Const ERROR_SSH_COMPRESSION_ERROR As Long = &H6A                     'Compression or decompression error.
Public Const ERROR_SSH_SERVICE_NOT_AVAILABLE As Long = &H6B                 'Service (sftp, shell, etc.) is not available.
Public Const ERROR_SSH_PROTOCOL_VERSION_NOT_SUPPORTED As Long = &H6C        'Version is not supported.
Public Const ERROR_SSH_HOST_KEY_NOT_VERIFIABLE As Long = &H6D               'Server key can not be verified.
Public Const ERROR_SSH_CONNECTION_LOST As Long = &H6E                       'Connection was lost by some reason.
Public Const ERROR_SSH_APPLICATION_CLOSED As Long = &H6F                    'User on the other side of connection closed application that led to disconnection.
Public Const ERROR_SSH_TOO_MANY_CONNECTIONS As Long = &H70                  'The server is overladen.
Public Const ERROR_SSH_AUTH_CANCELLED_BY_USER As Long = &H71                'User tired of invalid password entering.
Public Const ERROR_SSH_NO_MORE_AUTH_METHODS_AVAILABLE As Long = &H72        'There are no more methods for user authentication.
Public Const ERROR_SSH_ILLEGAL_USERNAME As Long = &H73                      'There is no user with specified username on the server.
Public Const ERROR_SSH_INTERNAL_ERROR As Long = &HC8                        'Internal error of implementation.
Public Const ERROR_SSH_NOT_CONNECTED As Long = &HDE                         'There is no connection but user tries to send data.
Public Const ERROR_SSH_CONNECTION_CANCELLED_BY_USER As Long = &H1F5         'The connection was cancelled by user.
Public Const ERROR_SSH_FORWARD_DISALLOWED As Long = &H1F6                   'SSH forward disallowed.
Public Const ERROR_SSH_ONKEYVALIDATE_NOT_ASSIGNED As Long = &H1F7           'The event handler for OnKeyValidate event, has not been specified by the application.
Public Const ERROR_SSH_TCP_CONNECTION_FAILED As Long = &H6001               'TCP connection failed.
Public Const ERROR_SSH_TCP_BIND_FAILED As Long = &H6002                     'TCP bind failed.