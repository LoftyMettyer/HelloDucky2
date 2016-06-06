Option Explicit On
Option Strict On

Namespace Classes.Workspace
   Public Enum FeedParticipantResponseCodeEnum
      Success = 1
      NetworkIssues = 2
      ConfigurationIssue = 3
      CodeException = 4
      InvalidUser = 5
      BusinessException = 6
   End Enum
End Namespace