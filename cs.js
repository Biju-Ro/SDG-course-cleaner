const fs = require('fs');
const XLSX = require('xlsx');

const workbook = XLSX.readFile('competency and skills.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(worksheet);

const competencyCategories = {
  'Systems Thinking': 'Systems Thinking (complex problem solving, relationship-focused)',
  'Life-cycle Thinking': 'Life-cycle Thinking (LCA, LCCA, whole-life cost analysis)',
  'Future Thinking': 'Future Thinking (scenario planning, forecasting / back-casting, trends and drivers)',
  'Interpersonal Competency': 'Interpersonal Competency (communications, mediation, stakeholder engagement)',
  'Strategic Thinking': 'Strategic Thinking (action planning, goal orientation, innovation)',
  'Mental Models': 'Mental Models (social norms, heuristics, biases)'
};

function cleanValue(value) {
  if (!value || value.trim() === '' || value.trim() === ' ') return 'Unavailable';
  return value.trim();
}

function categorizeCompetency(competencyText) {
  if (competencyText === 'Unavailable') return 'Unavailable';
  
  const text = competencyText.toLowerCase();
  if (text.includes('systems thinking') || text.includes('complex problem solving') || text.includes('relationship-focused')) {
    return 'Systems Thinking';
  }
  if (text.includes('life-cycle thinking') || text.includes('lca') || text.includes('lcca') || text.includes('whole-life cost')) {
    return 'Life-cycle Thinking';
  }
  if (text.includes('future thinking') || text.includes('scenario planning') || text.includes('forecasting') || text.includes('back-casting') || text.includes('trends and drivers')) {
    return 'Future Thinking';
  }
  if (text.includes('interpersonal competency') || text.includes('communications') || text.includes('mediation') || text.includes('stakeholder engagement')) {
    return 'Interpersonal Competency';
  }
  if (text.includes('strategic thinking') || text.includes('action planning') || text.includes('goal orientation') || text.includes('innovation')) {
    return 'Strategic Thinking';
  }
  if (text.includes('mental models') || text.includes('social norms') || text.includes('heuristics') || text.includes('biases')) {
    return 'Mental Models';
  }
  return 'Other';
}

function processData(data) {
  const normalizedData = [];
  const courseStats = new Map();
  const competencyStats = new Map();
  const categoryStats = new Map();
  const professorStats = new Map();
  
  Object.keys(competencyCategories).forEach(category => {
    categoryStats.set(category, 0);
  });
  categoryStats.set('Other', 0);

  data.forEach(row => {
    const courseName = cleanValue(row['course name']);
    if (courseName === 'Unavailable') return;

    const professorEmail = cleanValue(row['Professor email']);
    const professorName = cleanValue(row['Professor name']);

    let courseCompetencyCount = 0;
    const courseCompetencies = [];
    const courseCategories = new Set();

    for (let i = 1; i <= 6; i++) {
      const competencyValue = cleanValue(row[`competency and skills ${i}`]);
      if (competencyValue !== 'Unavailable') {
        courseCompetencyCount++;
        const category = categorizeCompetency(competencyValue);
        courseCategories.add(category);
        
        if (!competencyStats.has(competencyValue)) {
          competencyStats.set(competencyValue, 0);
        }
        competencyStats.set(competencyValue, competencyStats.get(competencyValue) + 1);
        
        categoryStats.set(category, categoryStats.get(category) + 1);
        
        normalizedData.push({
          courseName: courseName,
          competencyNumber: i,
          competencyText: competencyValue,
          competencyCategory: category,
          
          courseCompetencyCount: 0,
          courseCategoryCount: 0,
          competencyFrequency: 0,
          categoryFrequency: 0,
          
          professorName: professorName,
          professorEmail: professorEmail
        });
      }
    }

    courseStats.set(courseName, {
      competencyCount: courseCompetencyCount,
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
          totalCompetencies: 0,
          uniqueCategories: new Set(),
          avgCompetenciesPerCourse: 0
        });
      }
      const prof = professorStats.get(professorEmail);
      prof.courses++;
      prof.totalCompetencies += courseCompetencyCount;
      courseCategories.forEach(cat => prof.uniqueCategories.add(cat));
      prof.avgCompetenciesPerCourse = prof.totalCompetencies / prof.courses;
    }
  });

  normalizedData.forEach(row => {
    const courseInfo = courseStats.get(row.courseName);
    row.courseCompetencyCount = courseInfo.competencyCount;
    row.courseCategoryCount = courseInfo.categoryCount;
    row.competencyFrequency = competencyStats.get(row.competencyText);
    row.categoryFrequency = categoryStats.get(row.competencyCategory);
  });

  return { normalizedData, courseStats, competencyStats, categoryStats, professorStats };
}

function createSummaryData(normalizedData, courseStats, competencyStats, categoryStats, professorStats) {
  const totalCourses = courseStats.size;
  const totalMappings = normalizedData.length;
  
  const competencyFrequencyData = [];
  competencyStats.forEach((frequency, competency) => {
    competencyFrequencyData.push({
      competency: competency,
      category: categorizeCompetency(competency),
      frequency: frequency,
      percentageOfCourses: ((frequency / totalCourses) * 100).toFixed(2)
    });
  });
  competencyFrequencyData.sort((a, b) => b.frequency - a.frequency);

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
      competencyCount: stats.competencyCount,
      categoryCount: stats.categoryCount,
      categories: stats.categories.join(', '),
      complexity: stats.competencyCount > 4 ? 'High' : stats.competencyCount > 2 ? 'Medium' : stats.competencyCount > 0 ? 'Low' : 'None',
      diversity: stats.categoryCount > 4 ? 'High' : stats.categoryCount > 2 ? 'Medium' : stats.categoryCount > 0 ? 'Low' : 'None',
      professorName: stats.professor.name,
      professorEmail: stats.professor.email
    });
  });
  courseComplexityData.sort((a, b) => b.competencyCount - a.competencyCount);

  const professorAnalysisData = [];
  professorStats.forEach((stats, email) => {
    professorAnalysisData.push({
      professorEmail: email,
      professorName: stats.name,
      courseCount: stats.courses,
      totalCompetencies: stats.totalCompetencies,
      avgCompetenciesPerCourse: parseFloat(stats.avgCompetenciesPerCourse.toFixed(2)),
      uniqueCategories: stats.uniqueCategories.size,
      categoryDiversity: Array.from(stats.uniqueCategories).join(', ')
    });
  });
  professorAnalysisData.sort((a, b) => b.avgCompetenciesPerCourse - a.avgCompetenciesPerCourse);

  const competencyDistribution = {
    coursesWithNoCompetencies: totalCourses - new Set(normalizedData.map(row => row.courseName)).size,
    coursesWithCompetencies: new Set(normalizedData.map(row => row.courseName)).size,
    averageCompetenciesPerCourse: (totalMappings / new Set(normalizedData.map(row => row.courseName)).size).toFixed(2),
    totalCourses,
    totalMappings,
    uniqueCompetencies: competencyStats.size,
    activeCategories: categoryFrequencyData.length
  };

  return { competencyFrequencyData, categoryFrequencyData, courseComplexityData, professorAnalysisData, competencyDistribution };
}

function exportToExcel(normalizedData, summaryData) {
  const workbook = XLSX.utils.book_new();

  const mainSheet = XLSX.utils.json_to_sheet(normalizedData.map(row => ({
    'Course Name': row.courseName,
    'Competency Number': row.competencyNumber,
    'Competency Text': row.competencyText,
    'Competency Category': row.competencyCategory,
    'Course Competency Count': row.courseCompetencyCount,
    'Course Category Count': row.courseCategoryCount,
    'Competency Frequency': row.competencyFrequency,
    'Category Frequency': row.categoryFrequency,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const competencyAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.competencyFrequencyData.map(row => ({
    'Competency': row.competency,
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
    'Competency Count': row.competencyCount,
    'Category Count': row.categoryCount,
    'Categories': row.categories,
    'Competency Complexity': row.complexity,
    'Category Diversity': row.diversity,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const professorAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.professorAnalysisData.map(row => ({
    'Professor Email': row.professorEmail,
    'Professor Name': row.professorName,
    'Course Count': row.courseCount,
    'Total Competencies': row.totalCompetencies,
    'Avg Competencies per Course': row.avgCompetenciesPerCourse,
    'Unique Categories': row.uniqueCategories,
    'Category Diversity': row.categoryDiversity
  })));

  const overviewSheet = XLSX.utils.json_to_sheet([
    { Metric: 'Total Courses', Value: summaryData.competencyDistribution.totalCourses },
    { Metric: 'Courses with Competencies', Value: summaryData.competencyDistribution.coursesWithCompetencies },
    { Metric: 'Courses without Competencies', Value: summaryData.competencyDistribution.coursesWithNoCompetencies },
    { Metric: 'Total Course-Competency Mappings', Value: summaryData.competencyDistribution.totalMappings },
    { Metric: 'Average Competencies per Course', Value: summaryData.competencyDistribution.averageCompetenciesPerCourse },
    { Metric: 'Unique Competencies', Value: summaryData.competencyDistribution.uniqueCompetencies },
    { Metric: 'Active Categories', Value: summaryData.competencyDistribution.activeCategories },
    { Metric: 'Most Common Competency', Value: summaryData.competencyFrequencyData[0].competency },
    { Metric: 'Most Common Competency Frequency', Value: summaryData.competencyFrequencyData[0].frequency },
    { Metric: 'Most Common Category', Value: summaryData.categoryFrequencyData[0].category },
    { Metric: 'Most Common Category Frequency', Value: summaryData.categoryFrequencyData[0].frequency }
  ]);

  XLSX.utils.book_append_sheet(workbook, mainSheet, 'Course-Competency Data');
  XLSX.utils.book_append_sheet(workbook, competencyAnalysisSheet, 'Competency Analysis');
  XLSX.utils.book_append_sheet(workbook, categoryAnalysisSheet, 'Category Analysis');
  XLSX.utils.book_append_sheet(workbook, courseAnalysisSheet, 'Course Analysis');
  XLSX.utils.book_append_sheet(workbook, professorAnalysisSheet, 'Professor Analysis');
  XLSX.utils.book_append_sheet(workbook, overviewSheet, 'Overview');

  XLSX.writeFile(workbook, 'sustainability_competency_analysis.xlsx');
}

try {
  console.log('Processing sustainability competency data for Excel analysis...');
  
  const { normalizedData, courseStats, competencyStats, categoryStats, professorStats } = processData(rawData);
  const summaryData = createSummaryData(normalizedData, courseStats, competencyStats, categoryStats, professorStats);
  
  exportToExcel(normalizedData, summaryData);
  
  console.log('\n📊 Analysis Complete!');
  console.log('📁 File created: sustainability_competency_analysis.xlsx');
  console.log('\n📋 Excel sheets included:');
  console.log('   1. Course-Competency Data - Main data with numeric columns for analysis');
  console.log('   2. Competency Analysis - Frequency and popularity of each competency');
  console.log('   3. Category Analysis - Frequency of competency categories');
  console.log('   4. Course Analysis - Course complexity and competency diversity');
  console.log('   5. Professor Analysis - Professor competency teaching patterns');
  console.log('   6. Overview - Summary statistics');
  
  console.log('\n🔢 Numeric columns for analysis:');
  console.log('   • Course Competency Count - How many competencies each course has');
  console.log('   • Course Category Count - How many different categories per course');
  console.log('   • Competency Frequency - How popular each competency is across all courses');
  console.log('   • Category Frequency - How popular each category is across all courses');
  
  console.log(`\n📈 Quick Stats:`);
  console.log(`   • Total courses: ${summaryData.competencyDistribution.totalCourses}`);
  console.log(`   • Total course-competency pairs: ${summaryData.competencyDistribution.totalMappings}`);
  console.log(`   • Average competencies per course: ${summaryData.competencyDistribution.averageCompetenciesPerCourse}`);
  console.log(`   • Unique competencies: ${summaryData.competencyDistribution.uniqueCompetencies}`);
  console.log(`   • Most popular competency: ${summaryData.competencyFrequencyData[0].competency} (${summaryData.competencyFrequencyData[0].frequency} courses)`);
  console.log(`   • Most popular category: ${summaryData.categoryFrequencyData[0].category} (${summaryData.categoryFrequencyData[0].frequency} instances)`);

} catch (error) {
  console.error('❌ Error processing data:', error);
}