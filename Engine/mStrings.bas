Attribute VB_Name = "mStrings"
Option Explicit
Public Const e_strUnableToAuth = " àÏÑÝÈÖßt¹·tÚÊÚ×ÖÜ”ÀÐÏ¸—ËÑÌÈÝç¹×È—ÑÍÜ‡ÎÕ½ÍÓ¸—ÝÛ‰ÈÝè¼Ðà½ñÎŒØÙˆÜµÔŽ¶ÜÎÚ‰ËÑçµÃÚ¹Û—ŒÂÖÝ”·ÂÜtëÛå‰Û×”µÖâ¼—ÊÓÊÐÖ”ÉÔ×ÂÞ‰àÑÌˆÚÃÍÚÃîÒÚÐ‡ËãÁÎÏÂÛ£Œ‹–ÔètÂãÈß‰ßÎÙÞÙÆŸv£‰Ž˜ÓÜ”µÖâ¼—ÜÑÛÝÍæt“tæÛŒ‹–ÔètÂãÈß‰ßÎÙÞÙÆ¡v¥‰¿ØÙÚítÇÝÆ—ÝÔÎ‡Ñâ·ÐÜÊàÎÚÒÌÖ×¹"
Public Const e_strBanned = "­æÞŒÊÙÍ”–¢¼¢¼­ŒÏÙ×átÖá½åÐŒµÐÎÙÈÂÜ¿—Áµ•‡ÜÜ½ÔŽ·æÞØÍ‡ÊÙtÇÝÆ—ÊŒ×ÜÕÖ¹ÓŽÃÝ‰ÞÎÈÛãÂÔŽÈßÊà‰à×étÔÖÃìÕÐ‰×Úã¶ÂÐÀð‰ÎÎ‡ÉëµÓÓtæÏš‰º×æÆÚŽºæÛŒÝÏÍ”½ÏÑÃåßÕÎÕÍâ·Æœ"
Public Const e_strInquire = "¨æ‰Õ×ØÝÝÆÆŽµÙØáÝ‡ÜÜ¹áÈØÝÑ‰ÖÎ”ÈÉ×Ç—ËÍ×“ˆäÀÆÏÇÜ‰ÜØÚÜ”ÈÉÓtÝØØÕÖßÝÂÈŽÈæ‰ÛÞÙˆÇÉÑÞÃéÝŒÏÖÚéÁÔŽµë‰ÔÝÛØ®ƒåËî—ØÒÍÍèµÏÙÌà—ÏØÔˆ®t"
Public Const e_strCopy = "—æÙå‰ÈÖØt±ÏÇëÎŒ½ÏÑçŽ"
Public Function g_String(bStrText As String) As String
    g_String = mCrypt.Decrypt(bStrText)
End Function

