-- use sponsoriq;

-- Validateing the Data

SELECT COUNT(*) FROM Brand_Master;

SELECT COUNT(*) FROM Event_Master;

SELECT COUNT(*) FROM Sponsorship_History;

SELECT COUNT(*) FROM Decision_Makers;

SELECT COUNT(*) FROM Sales_CRM;

SELECT COUNT(*) FROM Marketing_Campaigns;

SELECT COUNT(*) FROM Industry_Benchmark;

SELECT COUNT(*) FROM Calendar;

SELECT COUNT(*) FROM Brand_IP_Fit;

-- 1. How many brands are there in our database?
select count(*) as Total_Brands from brand_master ;

-- 2. List all brands with their industry.
select Brand_name,Industry from brand_master;

-- Which are the Top 10 highest revenue brands?
select Brand_name,Annual_Revenue_Cr
from brand_master
order by Annual_Revenue_Cr desc
limit 10;

-- 4. Which brands spend the most on marketing?
select Brand_Name,Marketing_spend_Cr
from brand_master
order by Marketing_spend_Cr desc
limit 10;

-- 5. Which brands have the highest sponsorship budget?
select Brand_name,Sponsorship_Budget_Cr
from brand_master
order by Sponsorship_Budget_Cr desc
limit 10;

-- 6. Which industries generate the highest revenue?
select Industry,sum(Annual_Revenue_Cr) as Total_Revenue
from brand_master
group by  Industry
order by Total_Revenue desc;

-- 7. What is the average marketing spend by industry?
select Industry,round(avg(Marketing_Spend_Cr),2) as Marketing_Spend
from brand_master
group by Industry
order by Marketing_Spend;

-- 8. Which brands have above-average growth?
SELECT Brand_Name,
       Growth_Percentage
FROM Brand_Master
WHERE Growth_Percentage >
(
SELECT AVG(Growth_Percentage)
FROM Brand_Master
)
ORDER BY Growth_Percentage DESC;

-- 9. Which brands are listed on the stock market?
select Brand_Name,Stock_Listed
from Brand_Master
where Stock_Listed ='Yes';

-- 10. How many brands belong to each industry?
select Industry,count(Brand_Name) as Brands
From brand_master
group by Industry
order by Brands;

-- 11. Which events have the highest sponsorship cost?
select Event_name,Sponsorship_Cost_Cr
from event_master
order by Sponsorship_Cost_Cr desc
limit 5;

-- 12. Which events attract the largest audience?
select Event_Name,Annual_Footfall
from event_master
order by Annual_Footfall desc
limit 5;

-- 13. Which event categories are most common?
SELECT Event_Category,
       COUNT(*) AS Total_Events
FROM Event_Master
GROUP BY Event_Category
ORDER BY Total_Events DESC;

-- 14. Which brands invested the most in sponsorships?
SELECT b.Brand_name,s.Brand_ID,
       SUM(s.Investment_Cr) AS Total_Investment
FROM Sponsorship_History s
inner join brand_master b on b.Brand_ID=s.Brand_ID
GROUP BY b.Brand_ID
ORDER BY Total_Investment DESC;

-- 15. Which sponsorship type is most popular?
SELECT Sponsorship_Type,
       COUNT(*) AS Total
FROM Sponsorship_History
GROUP BY Sponsorship_Type
ORDER BY Total DESC;

-- 16. What is the average ROI by sponsorship type?
select Sponsorship_Type, round(avg(Estimated_ROI),2) as ROI 
from sponsorship_history
group by sponsorship_Type
order by ROI;

-- 17. Which deals are currently in the Proposal stage?
select b.Brand_Name,
       s.Proposal_Status,
       s.Deal_Value_Cr
FROM sales_crm s 
left join brand_master b on b.Brand_ID=s.Brand_ID
WHERE s.Proposal_Status='Proposal';

-- 18. Which deals have the highest probability of closing?
SELECT Brand_ID,
       Probability
FROM Sales_CRM
ORDER BY Probability DESC;

-- 19. Which marketing campaigns generated the highest ROI?
SELECT Campaign_Name,
       ROI
FROM Marketing_Campaigns
ORDER BY ROI DESC;

-- 20. Which brands have the best Brand-IP Match Score?
SELECT Brand_ID,
       Overall_Match_Score
FROM Brand_IP_Fit
ORDER BY Overall_Match_Score DESC;

-- 21. Which events are recommended most frequently?
SELECT Event_ID,
       COUNT(*) AS Recommendation_Count
FROM Brand_IP_Fit
GROUP BY Event_ID
ORDER BY Recommendation_Count DESC;

-- 22. Which industries have the highest average match score?
SELECT Industry_Fit,
       ROUND(AVG(Overall_Match_Score),2) AS Avg_Score
FROM Brand_IP_Fit
GROUP BY Industry_Fit
ORDER BY Avg_Score DESC;

-- 23. Which brands have multiple decision makers?
SELECT Brand_ID,
       COUNT(*) AS Decision_Makers
FROM Decision_Makers
GROUP BY Brand_ID
HAVING COUNT(*)>1;

-- 24. Which brands have active campaigns but no sponsorship history?
SELECT DISTINCT b.Brand_Name
FROM Brand_Master b
JOIN Marketing_Campaigns m
ON b.Brand_ID=m.Brand_ID
LEFT JOIN Sponsorship_History s
ON b.Brand_ID=s.Brand_ID
WHERE s.Brand_ID IS NULL;

-- 25. Which brands are the best sponsorship opportunities?
SELECT b.Brand_Name,
       b.Industry,
       b.Marketing_Spend_Cr,
       b.Sponsorship_Budget_Cr,
       f.Overall_Match_Score
FROM Brand_Master b
JOIN Brand_IP_Fit f
ON b.Brand_ID=f.Brand_ID
ORDER BY f.Overall_Match_Score DESC
LIMIT 20;

-- 26. Find the Top 5 brands by sponsorship investment with their event details.
SELECT
    b.Brand_Name,
    e.Event_Name,
    SUM(s.Investment_Cr) AS Total_Investment
FROM Sponsorship_History s
JOIN Brand_Master b
    ON s.Brand_ID = b.Brand_ID
JOIN Event_Master e
    ON s.Event_ID = e.Event_ID
GROUP BY b.Brand_Name, e.Event_Name
ORDER BY Total_Investment DESC
LIMIT 5;