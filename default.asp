<% 
	dim merchantId, agreementId, orderId, amount, curr, continueUrl, cancelUrl, callbackUrl, language, autocapture, paymentMethods, publicKey

	merchantId = 1
	agreementId = 1
	orderId = "1234"
	amount = 100
	curr = "DKK"
	continueUrl = "http://www.webshop.dk/completed/"
	cancelUrl = "http://www.webshop.dk/cancel/"
	callbackUrl = "http://www.webshop.dk/callback/"
	language = "da"
	autocapture = 0
	paymentMethods = "dankort"
	publicKey = "publickey"

	set payment = createObject("QuickPayV10.ActiveX.Payment")
	payment.MerchantId = merchantId
	payment.AgreementId = agreementId
	payment.OrderId = orderId
	payment.Amount = amount
	payment.Currency = curr
	payment.ContinueUrl = continueUrl
	payment.CancelUrl = cancelUrl
	payment.CallbackUrl = callbackUrl
	payment.Language = language
	payment.Autocapture = autocapture
	payment.PaymentMethods = paymentMethods
	payment.PublicKey = publicKey
	
	dim checksum
	checksum = payment.ComputeChecksum()	
%>

<!DOCTYPE html>
<html>
<body>
<form method="POST" action="https://payment.quickpay.net">
<input type="hidden" name="version" value="v10">
<input type="hidden" name="merchant_id" value="<%=merchantId %>">
<input type="hidden" name="agreement_id" value="<%=agreementId %>">
<input type="hidden" name="order_id" value="<%=orderId %>">
<input type="hidden" name="amount" value="<%=amount %>">
<input type="hidden" name="currency" value="<%=curr %>">
<input type="hidden" name="continueurl" value="<%=continueUrl %>">
<input type="hidden" name="cancelurl" value="<%=cancelUrl %>">
<input type="hidden" name="callbackurl" value="<%=callbackUrl %>">
<input type="hidden" name="language" value="<%=language %>">
<input type="hidden" name="autocapture" value="<%=autocapture %>">
<input type="hidden" name="paymentmethods" value="<%=paymentMethods %>">
<input type="hidden" name="checksum" value="<%=checksum %>">
<input type="submit" value="Continue to payment...">
</form>
</body>
</html>