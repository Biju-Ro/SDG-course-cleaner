const fs = require('fs');
const XLSX = require('xlsx');

const workbook = XLSX.readFile('keywords.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(worksheet);

const keywordCategories = {
  'Human Thriving': 'Human Thriving (appreciation, gratitude, inclusion, and belonging)',
  'Ethics': 'Ethics (responsibility, empathy, the social good)',
  'Justice': 'Justice (environmental, social, climate, human rights, diversity)',
  'Preservation': 'Preservation for Future Generations (ecological well-being)',
  'Equity': 'Equity (poverty reduction, women\'s rights, gender equality)',
  'Physical Well-being': 'Physical Well-being (mental health, social acceptance'
};

function cleanValue(value) {
  if (!value || value.trim() === '' || value.trim() === ' ') return 'Unavailable';
  return value.trim();
}

function categorizeKeyword(keyword) {
  if (keyword === 'Unavailable') return 'Unavailable';
  
  const keywordLower = keyword.toLowerCase();
  if (keywordLower.includes('human thriving') || keywordLower.includes('appreciation') || keywordLower.includes('gratitude') || keywordLower.includes('inclusion') || keywordLower.includes('belonging')) {
    return 'Human Thriving';
  }
  if (keywordLower.includes('ethics') || keywordLower.includes('responsibility') || keywordLower.includes('empathy') || keywordLower.includes('social good')) {
    return 'Ethics';
  }
  if (keywordLower.includes('justice') || keywordLower.includes('environmental') || keywordLower.includes('climate') || keywordLower.includes('human rights') || keywordLower.includes('diversity')) {
    return 'Justice';
  }
  if (keywordLower.includes('preservation') || keywordLower.includes('future generations') || keywordLower.includes('ecological')) {
    return 'Preservation';
  }
  if (keywordLower.includes('equity') || keywordLower.includes('poverty') || keywordLower.includes('women') || keywordLower.includes('gender equality')) {
    return 'Equity';
  }
  if (keywordLower.includes('physical') || keywordLower.includes('mental health') || keywordLower.includes('social acceptance') || keywordLower.includes('well-being')) {
    return 'Physical Well-being';
  }
  return 'Other';
}

function processData(data) {
  const normalizedData = [];
  const courseStats = new Map();
  const keywordStats = new Map();
  const categoryStats = new Map();
  
  Object.keys(keywordCategories).forEach(category => {
    categoryStats.set(category, 0);
  });
  categoryStats.set('Other', 0);

  data.forEach(row => {
    const courseName = cleanValue(row['course name']);
    if (courseName === 'Unavailable') return;

    const professorEmail = cleanValue(row['Professor email']) !== 'Unavailable' ? cleanValue(row['Professor email']) : cleanValue(row['Professor email_1']);
    const professorName = cleanValue(row['Professor name']) !== 'Unavailable' ? cleanValue(row['Professor name']) : cleanValue(row['Professor name_1']);

    let courseKeywordCount = 0;
    const courseKeywords = [];
    const courseCategories = new Set();

    for (let i = 1; i <= 6; i++) {
      const keywordValue = cleanValue(row[`KEYWORD ${i}`]);
      if (keywordValue !== 'Unavailable') {
        courseKeywordCount++;
        const category = categorizeKeyword(keywordValue);
        courseCategories.add(category);
        
        if (!keywordStats.has(keywordValue)) {
          keywordStats.set(keywordValue, 0);
        }
        keywordStats.set(keywordValue, keywordStats.get(keywordValue) + 1);
        
        categoryStats.set(category, categoryStats.get(category) + 1);
        
        normalizedData.push({
          courseName: courseName,
          keywordNumber: i,
          keywordText: keywordValue,
          keywordCategory: category,
          
          courseKeywordCount: 0,
          courseCategoryCount: 0,
          keywordFrequency: 0,
          categoryFrequency: 0,
          
          professorName: professorName,
          professorEmail: professorEmail
        });
      }
    }

    courseStats.set(courseName, {
      keywordCount: courseKeywordCount,
      categoryCount: courseCategories.size,
      categories: Array.from(courseCategories),
      professor: {
        name: professorName,
        email: professorEmail
      }
    });
  });

  normalizedData.forEach(row => {
    const courseInfo = courseStats.get(row.courseName);
    row.courseKeywordCount = courseInfo.keywordCount;
    row.courseCategoryCount = courseInfo.categoryCount;
    row.keywordFrequency = keywordStats.get(row.keywordText);
    row.categoryFrequency = categoryStats.get(row.keywordCategory);
  });

  return { normalizedData, courseStats, keywordStats, categoryStats };
}

function createSummaryData(normalizedData, courseStats, keywordStats, categoryStats) {
  const totalCourses = courseStats.size;
  const totalMappings = normalizedData.length;
  
  const keywordFrequencyData = [];
  keywordStats.forEach((frequency, keyword) => {
    keywordFrequencyData.push({
      keyword: keyword,
      category: categorizeKeyword(keyword),
      frequency: frequency,
      percentageOfCourses: ((frequency / totalCourses) * 100).toFixed(2)
    });
  });
  keywordFrequencyData.sort((a, b) => b.frequency - a.frequency);

  const categoryFrequencyData = [];
  categoryStats.forEach((frequency, category) => {
    if (frequency > 0) {
      categoryFrequencyData.push({
        category: category,
        frequency: frequency,
        percentageOfCourses: ((frequency / totalCourses) * 100).toFixed(2)
      });
    }
  });
  categoryFrequencyData.sort((a, b) => b.frequency - a.frequency);

  const courseComplexityData = [];
  courseStats.forEach((stats, courseName) => {
    courseComplexityData.push({
      courseName,
      keywordCount: stats.keywordCount,
      categoryCount: stats.categoryCount,
      categories: stats.categories.join(', '),
      complexity: stats.keywordCount > 4 ? 'High' : stats.keywordCount > 2 ? 'Medium' : 'Low',
      diversity: stats.categoryCount > 3 ? 'High' : stats.categoryCount > 1 ? 'Medium' : 'Low',
      professorName: stats.professor.name,
      professorEmail: stats.professor.email
    });
  });
  courseComplexityData.sort((a, b) => b.keywordCount - a.keywordCount);

  const keywordDistribution = {
    coursesWithNoKeywords: totalCourses - new Set(normalizedData.map(row => row.courseName)).size,
    coursesWithKeywords: new Set(normalizedData.map(row => row.courseName)).size,
    averageKeywordsPerCourse: (totalMappings / new Set(normalizedData.map(row => row.courseName)).size).toFixed(2),
    totalCourses,
    totalMappings,
    uniqueKeywords: keywordStats.size,
    activeCategories: categoryFrequencyData.length
  };

  return { keywordFrequencyData, categoryFrequencyData, courseComplexityData, keywordDistribution };
}

function exportToExcel(normalizedData, summaryData) {
  const workbook = XLSX.utils.book_new();

  const mainSheet = XLSX.utils.json_to_sheet(normalizedData.map(row => ({
    'Course Name': row.courseName,
    'Keyword Number': row.keywordNumber,
    'Keyword Text': row.keywordText,
    'Keyword Category': row.keywordCategory,
    'Course Keyword Count': row.courseKeywordCount,
    'Course Category Count': row.courseCategoryCount,
    'Keyword Frequency': row.keywordFrequency,
    'Category Frequency': row.categoryFrequency,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const keywordAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.keywordFrequencyData.map(row => ({
    'Keyword': row.keyword,
    'Category': row.category,
    'Frequency': row.frequency,
    'Percentage of Courses': row.percentageOfCourses
  })));

  const categoryAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.categoryFrequencyData.map(row => ({
    'Category': row.category,
    'Frequency': row.frequency,
    'Percentage of Courses': row.percentageOfCourses
  })));

  const courseAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.courseComplexityData.map(row => ({
    'Course Name': row.courseName,
    'Keyword Count': row.keywordCount,
    'Category Count': row.categoryCount,
    'Categories': row.categories,
    'Keyword Complexity': row.complexity,
    'Category Diversity': row.diversity,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const overviewSheet = XLSX.utils.json_to_sheet([
    { Metric: 'Total Courses', Value: summaryData.keywordDistribution.totalCourses },
    { Metric: 'Courses with Keywords', Value: summaryData.keywordDistribution.coursesWithKeywords },
    { Metric: 'Courses without Keywords', Value: summaryData.keywordDistribution.coursesWithNoKeywords },
    { Metric: 'Total Course-Keyword Mappings', Value: summaryData.keywordDistribution.totalMappings },
    { Metric: 'Average Keywords per Course', Value: summaryData.keywordDistribution.averageKeywordsPerCourse },
    { Metric: 'Unique Keywords', Value: summaryData.keywordDistribution.uniqueKeywords },
    { Metric: 'Active Categories', Value: summaryData.keywordDistribution.activeCategories },
    { Metric: 'Most Common Keyword', Value: summaryData.keywordFrequencyData[0].keyword },
    { Metric: 'Most Common Keyword Frequency', Value: summaryData.keywordFrequencyData[0].frequency },
    { Metric: 'Most Common Category', Value: summaryData.categoryFrequencyData[0].category },
    { Metric: 'Most Common Category Frequency', Value: summaryData.categoryFrequencyData[0].frequency }
  ]);

  XLSX.utils.book_append_sheet(workbook, mainSheet, 'Course-Keyword Data');
  XLSX.utils.book_append_sheet(workbook, keywordAnalysisSheet, 'Keyword Analysis');
  XLSX.utils.book_append_sheet(workbook, categoryAnalysisSheet, 'Category Analysis');
  XLSX.utils.book_append_sheet(workbook, courseAnalysisSheet, 'Course Analysis');
  XLSX.utils.book_append_sheet(workbook, overviewSheet, 'Overview');

  XLSX.writeFile(workbook, 'sustainability_keywords_analysis.xlsx');
}

try {
  console.log('Processing sustainability keywords data for Excel analysis...');
  
  const { normalizedData, courseStats, keywordStats, categoryStats } = processData(rawData);
  const summaryData = createSummaryData(normalizedData, courseStats, keywordStats, categoryStats);
  
  exportToExcel(normalizedData, summaryData);
  
  console.log('\n📊 Analysis Complete!');
  console.log('📁 File created: sustainability_keywords_analysis.xlsx');
  console.log('\n📋 Excel sheets included:');
  console.log('   1. Course-Keyword Data - Main data with numeric columns for analysis');
  console.log('   2. Keyword Analysis - Frequency and popularity of each keyword');
  console.log('   3. Category Analysis - Frequency of keyword categories');
  console.log('   4. Course Analysis - Course complexity and keyword diversity');
  console.log('   5. Overview - Summary statistics');
  
  console.log('\n🔢 Numeric columns for analysis:');
  console.log('   • Course Keyword Count - How many keywords each course has');
  console.log('   • Course Category Count - How many different categories per course');
  console.log('   • Keyword Frequency - How popular each keyword is across all courses');
  console.log('   • Category Frequency - How popular each category is across all courses');
  
  console.log(`\n📈 Quick Stats:`);
  console.log(`   • Total courses: ${summaryData.keywordDistribution.totalCourses}`);
  console.log(`   • Total course-keyword pairs: ${summaryData.keywordDistribution.totalMappings}`);
  console.log(`   • Average keywords per course: ${summaryData.keywordDistribution.averageKeywordsPerCourse}`);
  console.log(`   • Unique keywords: ${summaryData.keywordDistribution.uniqueKeywords}`);
  console.log(`   • Most popular keyword: ${summaryData.keywordFrequencyData[0].keyword} (${summaryData.keywordFrequencyData[0].frequency} courses)`);
  console.log(`   • Most popular category: ${summaryData.categoryFrequencyData[0].category} (${summaryData.categoryFrequencyData[0].frequency} instances)`);

} catch (error) {
  console.error('❌ Error processing data:', error);
}