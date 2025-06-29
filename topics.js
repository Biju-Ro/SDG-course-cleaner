const fs = require('fs');
const XLSX = require('xlsx');

const workbook = XLSX.readFile('sustainability_topics.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(worksheet);

const topicCategories = {
  'International Governance': 'International Governance (Sustainable Development Goals, treaties, collaboration, NGOs, IGOs)',
  'Local and National Governance': 'Local and National Governance (policies, incentives, accountability)',
  'Biosphere': 'Biosphere (the Anthropocene, ecological risks, biodiversity, ecology)',
  'Sustainable Food Systems': 'Sustainable Food Systems (organics, vertical & urban farming, aquaponics, hydroponics)',
  'Economics': 'Economics (natural capital, limits to growth, externalities, circular economy, ESG)',
  'Air and Climate Science': 'Air and Climate Science (air quality, GHG emissions)',
  'Decarbonization': 'Decarbonization (renewables, zero-carbon energy, carbon removal, sustainable transportation)',
  'Sustainable Design': 'Sustainable Design (products, materials, electronics, buildings, smart cities)',
  'Waste': 'Waste (Plastics, reuse, biodegradables, right-to-repair)'
};

function cleanValue(value) {
  if (!value || value.trim() === '' || value.trim() === ' ') return 'Unavailable';
  return value.trim();
}

function categorizeTopic(topicText) {
  if (topicText === 'Unavailable') return 'Unavailable';
  
  const text = topicText.toLowerCase();
  if (text.includes('international governance') || text.includes('sustainable development goals') || text.includes('treaties') || text.includes('ngos') || text.includes('igos')) {
    return 'International Governance';
  }
  if (text.includes('local and national governance') || text.includes('policies') || text.includes('incentives') || text.includes('accountability')) {
    return 'Local and National Governance';
  }
  if (text.includes('biosphere') || text.includes('anthropocene') || text.includes('ecological risks') || text.includes('biodiversity') || text.includes('ecology')) {
    return 'Biosphere';
  }
  if (text.includes('sustainable food systems') || text.includes('organics') || text.includes('vertical') || text.includes('urban farming') || text.includes('aquaponics') || text.includes('hydroponics')) {
    return 'Sustainable Food Systems';
  }
  if (text.includes('economics') || text.includes('natural capital') || text.includes('limits to growth') || text.includes('externalities') || text.includes('circular economy') || text.includes('esg')) {
    return 'Economics';
  }
  if (text.includes('air and climate science') || text.includes('air quality') || text.includes('ghg emissions')) {
    return 'Air and Climate Science';
  }
  if (text.includes('decarbonization') || text.includes('renewables') || text.includes('zero-carbon energy') || text.includes('carbon removal') || text.includes('sustainable transportation')) {
    return 'Decarbonization';
  }
  if (text.includes('sustainable design') || text.includes('products') || text.includes('materials') || text.includes('electronics') || text.includes('buildings') || text.includes('smart cities')) {
    return 'Sustainable Design';
  }
  if (text.includes('waste') || text.includes('plastics') || text.includes('reuse') || text.includes('biodegradables') || text.includes('right-to-repair')) {
    return 'Waste';
  }
  return 'Other';
}

function processData(data) {
  const normalizedData = [];
  const courseStats = new Map();
  const topicStats = new Map();
  const categoryStats = new Map();
  const professorStats = new Map();
  
  Object.keys(topicCategories).forEach(category => {
    categoryStats.set(category, 0);
  });
  categoryStats.set('Other', 0);

  data.forEach(row => {
    const courseName = cleanValue(row['course name']);
    if (courseName === 'Unavailable') return;

    const professorEmail = cleanValue(row['Professor email']);
    const professorName = cleanValue(row['Professor name']);

    let courseTopicCount = 0;
    const courseTopics = [];
    const courseCategories = new Set();

    for (let i = 1; i <= 9; i++) {
      const topicValue = cleanValue(row[`Sustainabilty Content Topic ${i}`]);
      if (topicValue !== 'Unavailable') {
        courseTopicCount++;
        const category = categorizeTopic(topicValue);
        courseCategories.add(category);
        
        if (!topicStats.has(topicValue)) {
          topicStats.set(topicValue, 0);
        }
        topicStats.set(topicValue, topicStats.get(topicValue) + 1);
        
        categoryStats.set(category, categoryStats.get(category) + 1);
        
        normalizedData.push({
          courseName: courseName,
          topicNumber: i,
          topicText: topicValue,
          topicCategory: category,
          
          courseTopicCount: 0,
          courseCategoryCount: 0,
          topicFrequency: 0,
          categoryFrequency: 0,
          
          professorName: professorName,
          professorEmail: professorEmail
        });
      }
    }

    courseStats.set(courseName, {
      topicCount: courseTopicCount,
      categoryCount: courseCategories.size,
      categories: Array.from(courseCategories),
      professor: {
        name: professorName,
        email: professorEmail
      }
    });

    if (professorEmail !== 'Unavailable') {
      if (!professorStats.has(professorEmail)) {
        professorStats.set(professorEmail, {
          name: professorName,
          courses: 0,
          totalTopics: 0,
          uniqueCategories: new Set(),
          avgTopicsPerCourse: 0
        });
      }
      const prof = professorStats.get(professorEmail);
      prof.courses++;
      prof.totalTopics += courseTopicCount;
      courseCategories.forEach(cat => prof.uniqueCategories.add(cat));
      prof.avgTopicsPerCourse = prof.totalTopics / prof.courses;
    }
  });

  normalizedData.forEach(row => {
    const courseInfo = courseStats.get(row.courseName);
    row.courseTopicCount = courseInfo.topicCount;
    row.courseCategoryCount = courseInfo.categoryCount;
    row.topicFrequency = topicStats.get(row.topicText);
    row.categoryFrequency = categoryStats.get(row.topicCategory);
  });

  return { normalizedData, courseStats, topicStats, categoryStats, professorStats };
}

function createSummaryData(normalizedData, courseStats, topicStats, categoryStats, professorStats) {
  const totalCourses = courseStats.size;
  const totalMappings = normalizedData.length;
  
  const topicFrequencyData = [];
  topicStats.forEach((frequency, topic) => {
    topicFrequencyData.push({
      topic: topic,
      category: categorizeTopic(topic),
      frequency: frequency,
      percentageOfCourses: ((frequency / totalCourses) * 100).toFixed(2)
    });
  });
  topicFrequencyData.sort((a, b) => b.frequency - a.frequency);

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
      topicCount: stats.topicCount,
      categoryCount: stats.categoryCount,
      categories: stats.categories.join(', '),
      complexity: stats.topicCount > 6 ? 'High' : stats.topicCount > 3 ? 'Medium' : stats.topicCount > 0 ? 'Low' : 'None',
      diversity: stats.categoryCount > 6 ? 'High' : stats.categoryCount > 3 ? 'Medium' : stats.categoryCount > 0 ? 'Low' : 'None',
      professorName: stats.professor.name,
      professorEmail: stats.professor.email
    });
  });
  courseComplexityData.sort((a, b) => b.topicCount - a.topicCount);

  const professorAnalysisData = [];
  professorStats.forEach((stats, email) => {
    professorAnalysisData.push({
      professorEmail: email,
      professorName: stats.name,
      courseCount: stats.courses,
      totalTopics: stats.totalTopics,
      avgTopicsPerCourse: parseFloat(stats.avgTopicsPerCourse.toFixed(2)),
      uniqueCategories: stats.uniqueCategories.size,
      categoryDiversity: Array.from(stats.uniqueCategories).join(', ')
    });
  });
  professorAnalysisData.sort((a, b) => b.avgTopicsPerCourse - a.avgTopicsPerCourse);

  const topicDistribution = {
    coursesWithNoTopics: totalCourses - new Set(normalizedData.map(row => row.courseName)).size,
    coursesWithTopics: new Set(normalizedData.map(row => row.courseName)).size,
    averageTopicsPerCourse: (totalMappings / new Set(normalizedData.map(row => row.courseName)).size).toFixed(2),
    totalCourses,
    totalMappings,
    uniqueTopics: topicStats.size,
    activeCategories: categoryFrequencyData.length
  };

  return { topicFrequencyData, categoryFrequencyData, courseComplexityData, professorAnalysisData, topicDistribution };
}

function exportToExcel(normalizedData, summaryData) {
  const workbook = XLSX.utils.book_new();

  const mainSheet = XLSX.utils.json_to_sheet(normalizedData.map(row => ({
    'Course Name': row.courseName,
    'Topic Number': row.topicNumber,
    'Topic Text': row.topicText,
    'Topic Category': row.topicCategory,
    'Course Topic Count': row.courseTopicCount,
    'Course Category Count': row.courseCategoryCount,
    'Topic Frequency': row.topicFrequency,
    'Category Frequency': row.categoryFrequency,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const topicAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.topicFrequencyData.map(row => ({
    'Topic': row.topic,
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
    'Topic Count': row.topicCount,
    'Category Count': row.categoryCount,
    'Categories': row.categories,
    'Topic Complexity': row.complexity,
    'Category Diversity': row.diversity,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const professorAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.professorAnalysisData.map(row => ({
    'Professor Email': row.professorEmail,
    'Professor Name': row.professorName,
    'Course Count': row.courseCount,
    'Total Topics': row.totalTopics,
    'Avg Topics per Course': row.avgTopicsPerCourse,
    'Unique Categories': row.uniqueCategories,
    'Category Diversity': row.categoryDiversity
  })));

  const overviewSheet = XLSX.utils.json_to_sheet([
    { Metric: 'Total Courses', Value: summaryData.topicDistribution.totalCourses },
    { Metric: 'Courses with Topics', Value: summaryData.topicDistribution.coursesWithTopics },
    { Metric: 'Courses without Topics', Value: summaryData.topicDistribution.coursesWithNoTopics },
    { Metric: 'Total Course-Topic Mappings', Value: summaryData.topicDistribution.totalMappings },
    { Metric: 'Average Topics per Course', Value: summaryData.topicDistribution.averageTopicsPerCourse },
    { Metric: 'Unique Topics', Value: summaryData.topicDistribution.uniqueTopics },
    { Metric: 'Active Categories', Value: summaryData.topicDistribution.activeCategories },
    { Metric: 'Most Common Topic', Value: summaryData.topicFrequencyData[0].topic },
    { Metric: 'Most Common Topic Frequency', Value: summaryData.topicFrequencyData[0].frequency },
    { Metric: 'Most Common Category', Value: summaryData.categoryFrequencyData[0].category },
    { Metric: 'Most Common Category Frequency', Value: summaryData.categoryFrequencyData[0].frequency }
  ]);

  XLSX.utils.book_append_sheet(workbook, mainSheet, 'Course-Topic Data');
  XLSX.utils.book_append_sheet(workbook, topicAnalysisSheet, 'Topic Analysis');
  XLSX.utils.book_append_sheet(workbook, categoryAnalysisSheet, 'Category Analysis');
  XLSX.utils.book_append_sheet(workbook, courseAnalysisSheet, 'Course Analysis');
  XLSX.utils.book_append_sheet(workbook, professorAnalysisSheet, 'Professor Analysis');
  XLSX.utils.book_append_sheet(workbook, overviewSheet, 'Overview');

  XLSX.writeFile(workbook, 'sustainability_topics_analysis.xlsx');
}

try {
  console.log('Processing sustainability topics data for Excel analysis...');
  
  const { normalizedData, courseStats, topicStats, categoryStats, professorStats } = processData(rawData);
  const summaryData = createSummaryData(normalizedData, courseStats, topicStats, categoryStats, professorStats);
  
  exportToExcel(normalizedData, summaryData);
  
  console.log('\n📊 Analysis Complete!');
  console.log('📁 File created: sustainability_topics_analysis.xlsx');
  console.log('\n📋 Excel sheets included:');
  console.log('   1. Course-Topic Data - Main data with numeric columns for analysis');
  console.log('   2. Topic Analysis - Frequency and popularity of each topic');
  console.log('   3. Category Analysis - Frequency of topic categories');
  console.log('   4. Course Analysis - Course complexity and topic diversity');
  console.log('   5. Professor Analysis - Professor topic teaching patterns');
  console.log('   6. Overview - Summary statistics');
  
  console.log('\n🔢 Numeric columns for analysis:');
  console.log('   • Course Topic Count - How many topics each course covers');
  console.log('   • Course Category Count - How many different categories per course');
  console.log('   • Topic Frequency - How popular each topic is across all courses');
  console.log('   • Category Frequency - How popular each category is across all courses');
  
  console.log(`\n📈 Quick Stats:`);
  console.log(`   • Total courses: ${summaryData.topicDistribution.totalCourses}`);
  console.log(`   • Total course-topic pairs: ${summaryData.topicDistribution.totalMappings}`);
  console.log(`   • Average topics per course: ${summaryData.topicDistribution.averageTopicsPerCourse}`);
  console.log(`   • Unique topics: ${summaryData.topicDistribution.uniqueTopics}`);
  console.log(`   • Most popular topic: ${summaryData.topicFrequencyData[0].topic} (${summaryData.topicFrequencyData[0].frequency} courses)`);
  console.log(`   • Most popular category: ${summaryData.categoryFrequencyData[0].category} (${summaryData.categoryFrequencyData[0].frequency} instances)`);

} catch (error) {
  console.error('❌ Error processing data:', error);
}