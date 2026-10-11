-- create database SponsorIQ;
-- use SponsorIQ;

-- 1. Brand Master
CREATE TABLE Brand_Master (
    Brand_ID VARCHAR(20) PRIMARY KEY,
    Brand_Name VARCHAR(150) NOT NULL,
    Parent_Company VARCHAR(150),
    Industry VARCHAR(100),
    Sub_Category VARCHAR(100),
    HQ_City VARCHAR(100),
    HQ_State VARCHAR(100),
    HQ_Country VARCHAR(100),
    Founded_Year INT,
    Annual_Revenue_Cr DECIMAL(15,2),
    Marketing_Spend_Cr DECIMAL(15,2),
    Sponsorship_Budget_Cr DECIMAL(15,2),
    Employee_Count INT,
    Website VARCHAR(255),
    Growth_Percentage DECIMAL(6,2),
    Brand_Type VARCHAR(20),
    Target_Audience VARCHAR(255),
    Global_or_Indian VARCHAR(20),
    Stock_Listed VARCHAR(10)
);

-- 2. Event Master

CREATE TABLE Event_Master (
    Event_ID VARCHAR(20) PRIMARY KEY,
    Event_Name VARCHAR(150),
    Event_Category VARCHAR(100),
    Organizer VARCHAR(150),
    Audience_Type VARCHAR(100),
    Annual_Footfall INT,
    Cities INT,
    Reach_Million DECIMAL(10,2),
    Sponsorship_Cost_Cr DECIMAL(12,2),
    Title_Sponsor VARCHAR(150),
    Associate_Sponsor VARCHAR(150),
    Season VARCHAR(50)
);

-- 3. Sponsorship History

CREATE TABLE Sponsorship_History (
    Sponsorship_ID VARCHAR(20) PRIMARY KEY,
    Brand_ID VARCHAR(20),
    Event_ID VARCHAR(20),
    Year INT,
    Sponsorship_Type VARCHAR(100),
    Investment_Cr DECIMAL(12,2),
    Duration VARCHAR(50),
    Status VARCHAR(50),
    Estimated_Reach DECIMAL(12,2),
    Estimated_ROI DECIMAL(8,2),

    FOREIGN KEY (Brand_ID) REFERENCES Brand_Master(Brand_ID),
    FOREIGN KEY (Event_ID) REFERENCES Event_Master(Event_ID)
);

-- 4. Decision Makers

CREATE TABLE Decision_Makers (
    Contact_ID VARCHAR(20) PRIMARY KEY,
    Brand_ID VARCHAR(20),
    Name VARCHAR(150),
    Designation VARCHAR(150),
    Department VARCHAR(100),
    LinkedIn_URL VARCHAR(255),
    Email VARCHAR(255),
    Phone VARCHAR(30),
    Contact_Status VARCHAR(50),
    Last_Contact_Date DATE,

    FOREIGN KEY (Brand_ID) REFERENCES Brand_Master(Brand_ID)
);

-- 5. Sales CRM

CREATE TABLE Sales_CRM (
    CRM_ID VARCHAR(20) PRIMARY KEY,
    Brand_ID VARCHAR(20),
    Proposal_Status VARCHAR(50),
    Proposal_Date DATE,
    Meeting_Date DATE,
    Follow_Up_Date DATE,
    Deal_Stage VARCHAR(100),
    Deal_Value_Cr DECIMAL(12,2),
    Probability INT,
    Remarks TEXT,

    FOREIGN KEY (Brand_ID) REFERENCES Brand_Master(Brand_ID)
);

-- 6. Marketing Campaigns    

CREATE TABLE Marketing_Campaigns (
    Campaign_ID VARCHAR(20) PRIMARY KEY,
    Brand_ID VARCHAR(20),
    Campaign_Name VARCHAR(150),
    Campaign_Type VARCHAR(100),
    Budget_Cr DECIMAL(12,2),
    Platform VARCHAR(100),
    Start_Date DATE,
    End_Date DATE,
    Reach DECIMAL(12,2),
    ROI DECIMAL(8,2),

    FOREIGN KEY (Brand_ID) REFERENCES Brand_Master(Brand_ID)
);

-- 7. Industry Benchmark

CREATE TABLE Industry_Benchmark (
    Industry VARCHAR(100) PRIMARY KEY,
    Avg_Revenue_Cr DECIMAL(15,2),
    Avg_Marketing_Spend_Cr DECIMAL(15,2),
    Avg_Sponsorship_Budget_Cr DECIMAL(15,2),
    Avg_Growth DECIMAL(6,2),
    Number_of_Brands INT
);

-- 8. Calendar

CREATE TABLE Calendar (
    Date DATE PRIMARY KEY,
    Day VARCHAR(20),
    Month VARCHAR(20),
    Quarter VARCHAR(20),
    Year INT,
    Financial_Year VARCHAR(20),
    Week_Number INT
);

-- 9. Brand IP Fit
CREATE TABLE Brand_IP_Fit (
    Fit_ID VARCHAR(20) PRIMARY KEY,
    Brand_ID VARCHAR(20),
    Event_ID VARCHAR(20),
    Industry_Fit INT,
    Audience_Fit INT,
    Budget_Fit INT,
    Geography_Fit INT,
    Brand_Value_Fit INT,
    Competition_Level VARCHAR(50),
    Overall_Match_Score DECIMAL(5,2),
    Recommended_Role VARCHAR(100),
    Confidence_Level VARCHAR(50),
    Recommendation_Reason TEXT,

    FOREIGN KEY (Brand_ID) REFERENCES Brand_Master(Brand_ID),
    FOREIGN KEY (Event_ID) REFERENCES Event_Master(Event_ID)
);

select * from brand_master; 
