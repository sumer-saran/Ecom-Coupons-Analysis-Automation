# Ecom-Coupons-Analysis-Automation
This is the code for Coupons Analysis Automation for An Orders Report Generated for Ecommerce Marketplace.

STEPS FOR USING THE AUTOMATION :
                             
1.	Download the OMS report, open it and Save As Data.xlsx in the Folder containing the python files.
2.	Three python executable files will be present named Step1, Step2 and Step3, run them in order.
3.	After running all three an Output.xlsx file will be generated

The OutputFile will have 2 sheets 
Sheet 1: Analysis
This will have all the coupons, no. of orders per coupon, revenue generated, average order value and percentage contribution to the total revenue.
I have sorted this sheet decreasingly by no. of orders per coupon.

Sheet 2: Comparison
This will have the Start Date of the respective coupons along with revenue generated a week before/after the first usage of the coupon. It also shows % increase/decrease in both cases. 
This sheet has been decreasingly sorted by % increase/decrease.

**NOTE** : This Code Is Valid For An Orders Report In A Particular Format, If You Want To Reuse It You Can Modify It By Changing The Column Names As Required For Your Orders Report.
