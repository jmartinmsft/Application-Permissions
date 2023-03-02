# Application-Permissions
 This PowerShell script leverages three separate Azure AD applications to demonstrate how application permissions affect access to mailboxes in Exchange Online.
 
 The first application should have API permissions assigned to it for Graph and EWS. No application access policy or RBAC permissions should be applied to the application.
 The second application should have API permissions assigned to it for Graph and EWS. An application access policy should be applied to the applicaion.
 The third application should have no API permissions assigned to it. RBAC permissions should be applied to this application.
 
 