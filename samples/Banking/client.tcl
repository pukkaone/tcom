package require tcom

set bank [::tcom::ref createobject "Banking.Bank"]
set account [$bank CreateAccount]
puts [$account Balance]
$account Deposit 20
puts [$account Balance]
$account Withdraw 10
puts [$account Balance]
