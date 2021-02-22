#include <Inet.au3>
#include <Json.au3>
$URL = "https://api.abucoins.com/products/stats"
$data = _INetGetSource($URL)
$object = json_decode($data)
Local $i = 0

While 1
    ;$product_id = Json_Get($object, '[' & $i & '].product_id')
    ;If @error Then ExitLoop
    ;If $product_id = "ETH-BTC" Then
        $volume_USD = Json_Get($object, '[' & $i & '].product_id')
        ConsoleWrite('$volume_USD = ' & $volume_USD & @CRLF ) ;### Debug Console
    ;EndIf
    $i = $i + 1
WEnd