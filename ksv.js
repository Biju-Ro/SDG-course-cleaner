const fs = require('fs');
const XLSX = require('xlsx');

const workbook = XLSX.readFile('KSV data.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(worksheet);

const sustainabilityFocusLevels = {
  'Sustainability-focused': 'Sustainability-focused (more than 75% of class instruction dedicated to sustainability content)',
  'Sustainability-related': 'Sustainability-related (more than 25% of class instruction dedicated to sustainability content)', 
  'Not a sustainability course': 'Not a sustainability course'
};

function cleanValue(value) {
  if (!value || value.trim() === '' || value.trim() === ' ') return 'Unavailable';
  return value.trim();
}

function yesNoToNumeric(value) {
  const cleaned = cleanValue(value);
  if (cleaned === 'Unavailable') return 0;
  
  const text = cleaned.toLowerCase();
  if (text === 'yes') return 1;
  if (text === 'no') return -1;
  if (text === 'not sure') return 0;
  return 0;
}

function percentageToNumeric(percentageText) {
  const cleaned = cleanValue(percentageText);
  if (cleaned === 'Unavailable') return 0;
  
  const text = cleaned.toLowerCase();
  if (text.includes('more than 75%') || text.includes('significant')) return 4;
  if (text.includes('between 25% to 75%') || text.includes('roughly half')) return 3;
  if (text.includes('less than 25%') || text.includes('some')) return 2;
  return 1;
}

function focusToNumeric(focusText) {
  const cleaned = cleanValue(focusText);
  if (cleaned === 'Unavailable') return 0;
  
  const text = cleaned.toLowerCase();
  if (text.includes('sustainability-focused') || text.includes('more than 75%')) return 3;
  if (text.includes('sustainability-related') || text.includes('more than 25%')) return 2;
  if (text.includes('not a sustainability course')) return 1;
  return 0;
}

function categorizeFocusLevel(focusText) {
  const cleaned = cleanValue(focusText);
  if (cleaned === 'Unavailable') return 'Unavailable';
  
  const text = cleaned.toLowerCase();
  if (text.includes('sustainability-focused')) return 'Sustainability-focused';
  if (text.includes('sustainability-related')) return 'Sustainability-related';
  if (text.includes('not a sustainability course')) return 'Not a sustainability course';
  return 'Other';
}

function processData(data) {
  const processedData = [];
  const ksvStats = {
    values: { yes: 0, no: 0, notSure: 0 },
    knowledge: { yes: 0, no: 0, notSure: 0 },
    skills: { yes: 0, no: 0, notSure: 0 }
  };
  const focusStats = new Map();
  const professorStats = new Map();
  
  Object.keys(sustainabilityFocusLevels).forEach(level => {
    focusStats.set(level, 0);
  });

  data.forEach(row => {
    const courseName = cleanValue(row['course name']);
    if (courseName === 'Unavailable') return;

    const sustainabilityValues = {
      includes: cleanValue(row['includes sustainability values?']),
      percentage: cleanValue(row['percentage sustainability values']),
      includesNumeric: yesNoToNumeric(row['includes sustainability values?']),
      percentageNumeric: percentageToNumeric(row['percentage sustainability values'])
    };

    const sustainabilityKnowledge = {
      includes: cleanValue(row['includes sustainability knowledge?']),
      percentage: cleanValue(row['percentage sustainability knowledge']),
      includesNumeric: yesNoToNumeric(row['includes sustainability knowledge?']),
      percentageNumeric: percentageToNumeric(row['percentage sustainability knowledge'])
    };

    const sustainabilitySkills = {
      includes: cleanValue(row['includes sustainability skills?']),
      percentage: cleanValue(row['percentage sustainability skills']),
      includesNumeric: yesNoToNumeric(row['includes sustainability skills?']),
      percentageNumeric: percentageToNumeric(row['percentage sustainability skills'])
    };

    const focusLevel = cleanValue(row['how sustainability focused is your class?']);
    const focusCategory = categorizeFocusLevel(focusLevel);
    const focusNumeric = focusToNumeric(focusLevel);

    const professorEmail = cleanValue(row['Professor email']);
    const professorName = cleanValue(row['Professor name']);

    const ksvScore = (
      sustainabilityValues.includesNumeric + sustainabilityValues.percentageNumeric +
      sustainabilityKnowledge.includesNumeric + sustainabilityKnowledge.percentageNumeric +
      sustainabilitySkills.includesNumeric + sustainabilitySkills.percentageNumeric
    ) / 6;

    const ksvCompleteness = [
      sustainabilityValues.includes !== 'Unavailable' ? 1 : 0,
      sustainabilityKnowledge.includes !== 'Unavailable' ? 1 : 0,
      sustainabilitySkills.includes !== 'Unavailable' ? 1 : 0
    ].reduce((a, b) => a + b, 0);

    const hasAllThreeKSV = (
      sustainabilityValues.includesNumeric > 0 &&
      sustainabilityKnowledge.includesNumeric > 0 &&
      sustainabilitySkills.includesNumeric > 0
    ) ? 1 : 0;

    if (sustainabilityValues.includes !== 'Unavailable') {
      const valuesResponse = sustainabilityValues.includes.toLowerCase();
      if (valuesResponse === 'yes') ksvStats.values.yes++;
      else if (valuesResponse === 'no') ksvStats.values.no++;
      else if (valuesResponse === 'not sure') ksvStats.values.notSure++;
    }

    if (sustainabilityKnowledge.includes !== 'Unavailable') {
      const knowledgeResponse = sustainabilityKnowledge.includes.toLowerCase();
      if (knowledgeResponse === 'yes') ksvStats.knowledge.yes++;
      else if (knowledgeResponse === 'no') ksvStats.knowledge.no++;
      else if (knowledgeResponse === 'not sure') ksvStats.knowledge.notSure++;
    }

    if (sustainabilitySkills.includes !== 'Unavailable') {
      const skillsResponse = sustainabilitySkills.includes.toLowerCase();
      if (skillsResponse === 'yes') ksvStats.skills.yes++;
      else if (skillsResponse === 'no') ksvStats.skills.no++;
      else if (skillsResponse === 'not sure') ksvStats.skills.notSure++;
    }

    if (focusCategory !== 'Unavailable' && focusCategory !== 'Other') {
      focusStats.set(focusCategory, focusStats.get(focusCategory) + 1);
    }

    if (professorEmail !== 'Unavailable') {
      if (!professorStats.has(professorEmail)) {
        professorStats.set(professorEmail, {
          name: professorName,
          courses: 0,
          avgKSVScore: 0,
          avgFocusScore: 0,
          totalKSVScore: 0,
          totalFocusScore: 0
        });
      }
      const prof = professorStats.get(professorEmail);
      prof.courses++;
      prof.totalKSVScore += ksvScore;
      prof.totalFocusScore += focusNumeric;
      prof.avgKSVScore = prof.totalKSVScore / prof.courses;
      prof.avgFocusScore = prof.totalFocusScore / prof.courses;
    }

    processedData.push({
      courseName,
      
      sustainabilityValuesIncludes: sustainabilityValues.includes,
      sustainabilityValuesPercentage: sustainabilityValues.percentage,
      sustainabilityValuesIncludesNum: sustainabilityValues.includesNumeric,
      sustainabilityValuesPercentageNum: sustainabilityValues.percentageNumeric,
      
      sustainabilityKnowledgeIncludes: sustainabilityKnowledge.includes,
      sustainabilityKnowledgePercentage: sustainabilityKnowledge.percentage,
      sustainabilityKnowledgeIncludesNum: sustainabilityKnowledge.includesNumeric,
      sustainabilityKnowledgePercentageNum: sustainabilityKnowledge.percentageNumeric,
      
      sustainabilitySkillsIncludes: sustainabilitySkills.includes,
      sustainabilitySkillsPercentage: sustainabilitySkills.percentage,
      sustainabilitySkillsIncludesNum: sustainabilitySkills.includesNumeric,
      sustainabilitySkillsPercentageNum: sustainabilitySkills.percentageNumeric,
      
      focusLevel: focusLevel,
      focusCategory: focusCategory,
      focusNumeric: focusNumeric,
      
      ksvScore: parseFloat(ksvScore.toFixed(2)),
      ksvCompleteness: ksvCompleteness,
      hasAllThreeKSV: hasAllThreeKSV,
      
      professorName: professorName,
      professorEmail: professorEmail
    });
  });

  return { processedData, ksvStats, focusStats, professorStats };
}

function createSummaryData(processedData, ksvStats, focusStats, professorStats) {
  const totalCourses = processedData.length;
  
  const ksvAnalysisData = [
    {
      component: 'Values',
      yesCount: ksvStats.values.yes,
      noCount: ksvStats.values.no,
      notSureCount: ksvStats.values.notSure,
      yesPercentage: ((ksvStats.values.yes / totalCourses) * 100).toFixed(2),
      positiveResponseRate: ((ksvStats.values.yes / (ksvStats.values.yes + ksvStats.values.no + ksvStats.values.notSure)) * 100).toFixed(2)
    },
    {
      component: 'Knowledge',
      yesCount: ksvStats.knowledge.yes,
      noCount: ksvStats.knowledge.no,
      notSureCount: ksvStats.knowledge.notSure,
      yesPercentage: ((ksvStats.knowledge.yes / totalCourses) * 100).toFixed(2),
      positiveResponseRate: ((ksvStats.knowledge.yes / (ksvStats.knowledge.yes + ksvStats.knowledge.no + ksvStats.knowledge.notSure)) * 100).toFixed(2)
    },
    {
      component: 'Skills',
      yesCount: ksvStats.skills.yes,
      noCount: ksvStats.skills.no,
      notSureCount: ksvStats.skills.notSure,
      yesPercentage: ((ksvStats.skills.yes / totalCourses) * 100).toFixed(2),
      positiveResponseRate: ((ksvStats.skills.yes / (ksvStats.skills.yes + ksvStats.skills.no + ksvStats.skills.notSure)) * 100).toFixed(2)
    }
  ];

  const focusAnalysisData = [];
  focusStats.forEach((count, level) => {
    focusAnalysisData.push({
      focusLevel: level,
      count: count,
      percentage: ((count / totalCourses) * 100).toFixed(2)
    });
  });
  focusAnalysisData.sort((a, b) => b.count - a.count);

  const professorAnalysisData = [];
  professorStats.forEach((stats, email) => {
    professorAnalysisData.push({
      professorEmail: email,
      professorName: stats.name,
      courseCount: stats.courses,
      avgKSVScore: parseFloat(stats.avgKSVScore.toFixed(2)),
      avgFocusScore: parseFloat(stats.avgFocusScore.toFixed(2))
    });
  });
  professorAnalysisData.sort((a, b) => b.avgKSVScore - a.avgKSVScore);

  const courseComplexityData = processedData
    .map(course => ({
      courseName: course.courseName,
      ksvScore: course.ksvScore,
      ksvCompleteness: course.ksvCompleteness,
      hasAllThreeKSV: course.hasAllThreeKSV,
      focusCategory: course.focusCategory,
      focusNumeric: course.focusNumeric,
      complexity: course.ksvScore > 2 ? 'High' : course.ksvScore > 1 ? 'Medium' : 'Low',
      professorName: course.professorName
    }))
    .sort((a, b) => b.ksvScore - a.ksvScore);

  const overviewStats = {
    totalCourses,
    coursesWithAllKSV: processedData.filter(c => c.hasAllThreeKSV === 1).length,
    averageKSVScore: (processedData.reduce((sum, c) => sum + c.ksvScore, 0) / totalCourses).toFixed(2),
    sustainabilityFocusedCourses: focusStats.get('Sustainability-focused') || 0,
    sustainabilityRelatedCourses: focusStats.get('Sustainability-related') || 0,
    nonSustainabilityCourses: focusStats.get('Not a sustainability course') || 0,
    uniqueProfessors: professorStats.size
  };

  return { ksvAnalysisData, focusAnalysisData, professorAnalysisData, courseComplexityData, overviewStats };
}

function exportToExcel(processedData, summaryData) {
  const workbook = XLSX.utils.book_new();

  const mainSheet = XLSX.utils.json_to_sheet(processedData.map(row => ({
    'Course Name': row.courseName,
    'Values (Text)': row.sustainabilityValuesIncludes,
    'Values (Numeric)': row.sustainabilityValuesIncludesNum,
    'Values % (Text)': row.sustainabilityValuesPercentage,
    'Values % (Numeric)': row.sustainabilityValuesPercentageNum,
    'Knowledge (Text)': row.sustainabilityKnowledgeIncludes,
    'Knowledge (Numeric)': row.sustainabilityKnowledgeIncludesNum,
    'Knowledge % (Text)': row.sustainabilityKnowledgePercentage,
    'Knowledge % (Numeric)': row.sustainabilityKnowledgePercentageNum,
    'Skills (Text)': row.sustainabilitySkillsIncludes,
    'Skills (Numeric)': row.sustainabilitySkillsIncludesNum,
    'Skills % (Text)': row.sustainabilitySkillsPercentage,
    'Skills % (Numeric)': row.sustainabilitySkillsPercentageNum,
    'Focus Level': row.focusLevel,
    'Focus Category': row.focusCategory,
    'Focus Numeric': row.focusNumeric,
    'KSV Score': row.ksvScore,
    'KSV Completeness': row.ksvCompleteness,
    'Has All Three KSV': row.hasAllThreeKSV,
    'Professor Name': row.professorName,
    'Professor Email': row.professorEmail
  })));

  const ksvAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.ksvAnalysisData.map(row => ({
    'KSV Component': row.component,
    'Yes Count': row.yesCount,
    'No Count': row.noCount,
    'Not Sure Count': row.notSureCount,
    'Yes Percentage': row.yesPercentage,
    'Positive Response Rate': row.positiveResponseRate
  })));

  const focusAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.focusAnalysisData.map(row => ({
    'Focus Level': row.focusLevel,
    'Count': row.count,
    'Percentage': row.percentage
  })));

  const professorAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.professorAnalysisData.map(row => ({
    'Professor Email': row.professorEmail,
    'Professor Name': row.professorName,
    'Course Count': row.courseCount,
    'Avg KSV Score': row.avgKSVScore,
    'Avg Focus Score': row.avgFocusScore
  })));

  const courseAnalysisSheet = XLSX.utils.json_to_sheet(summaryData.courseComplexityData.map(row => ({
    'Course Name': row.courseName,
    'KSV Score': row.ksvScore,
    'KSV Completeness': row.ksvCompleteness,
    'Has All Three KSV': row.hasAllThreeKSV,
    'Focus Category': row.focusCategory,
    'Focus Numeric': row.focusNumeric,
    'Complexity': row.complexity,
    'Professor Name': row.professorName
  })));

  const overviewSheet = XLSX.utils.json_to_sheet([
    { Metric: 'Total Courses', Value: summaryData.overviewStats.totalCourses },
    { Metric: 'Courses with All KSV Components', Value: summaryData.overviewStats.coursesWithAllKSV },
    { Metric: 'Average KSV Score', Value: summaryData.overviewStats.averageKSVScore },
    { Metric: 'Sustainability-focused Courses', Value: summaryData.overviewStats.sustainabilityFocusedCourses },
    { Metric: 'Sustainability-related Courses', Value: summaryData.overviewStats.sustainabilityRelatedCourses },
    { Metric: 'Non-sustainability Courses', Value: summaryData.overviewStats.nonSustainabilityCourses },
    { Metric: 'Unique Professors', Value: summaryData.overviewStats.uniqueProfessors },
    { Metric: 'Values Yes Rate', Value: `${summaryData.ksvAnalysisData[0].yesPercentage}%` },
    { Metric: 'Knowledge Yes Rate', Value: `${summaryData.ksvAnalysisData[1].yesPercentage}%` },
    { Metric: 'Skills Yes Rate', Value: `${summaryData.ksvAnalysisData[2].yesPercentage}%` }
  ]);

  XLSX.utils.book_append_sheet(workbook, mainSheet, 'KSV Course Data');
  XLSX.utils.book_append_sheet(workbook, ksvAnalysisSheet, 'KSV Component Analysis');
  XLSX.utils.book_append_sheet(workbook, focusAnalysisSheet, 'Focus Level Analysis');
  XLSX.utils.book_append_sheet(workbook, professorAnalysisSheet, 'Professor Analysis');
  XLSX.utils.book_append_sheet(workbook, courseAnalysisSheet, 'Course Analysis');
  XLSX.utils.book_append_sheet(workbook, overviewSheet, 'Overview');

  XLSX.writeFile(workbook, 'ksv_sustainability_analysis.xlsx');
}

try {
  console.log('Processing KSV sustainability data for Excel analysis...');
  
  const { processedData, ksvStats, focusStats, professorStats } = processData(rawData);
  const summaryData = createSummaryData(processedData, ksvStats, focusStats, professorStats);
  
  exportToExcel(processedData, summaryData);
  
  console.log('\n📊 Analysis Complete!');
  console.log('📁 File created: ksv_sustainability_analysis.xlsx');
  console.log('\n📋 Excel sheets included:');
  console.log('   1. KSV Course Data - Main data with numeric columns for analysis');
  console.log('   2. KSV Component Analysis - Values, Knowledge, Skills breakdown');
  console.log('   3. Focus Level Analysis - Sustainability focus distribution');
  console.log('   4. Professor Analysis - Professor performance metrics');
  console.log('   5. Course Analysis - Course complexity and completeness');
  console.log('   6. Overview - Summary statistics');
  
  console.log('\n🔢 Numeric columns for analysis:');
  console.log('   • Values/Knowledge/Skills (Numeric) - Yes(1)/No(-1)/Not Sure(0)');
  console.log('   • Values/Knowledge/Skills % (Numeric) - Percentage levels (0-4 scale)');
  console.log('   • Focus Numeric - Sustainability focus level (0-3 scale)');
  console.log('   • KSV Score - Overall KSV rating (-1 to 4 scale)');
  console.log('   • KSV Completeness - How many KSV components are filled (0-3)');
  console.log('   • Has All Three KSV - Binary indicator (0/1)');
  
  console.log(`\n📈 Quick Stats:`);
  console.log(`   • Total courses: ${summaryData.overviewStats.totalCourses}`);
  console.log(`   • Courses with all KSV components: ${summaryData.overviewStats.coursesWithAllKSV}`);
  console.log(`   • Average KSV score: ${summaryData.overviewStats.averageKSVScore}`);
  console.log(`   • Sustainability-focused courses: ${summaryData.overviewStats.sustainabilityFocusedCourses}`);
  console.log(`   • Sustainability-related courses: ${summaryData.overviewStats.sustainabilityRelatedCourses}`);
  console.log(`   • Values yes rate: ${summaryData.ksvAnalysisData[0].yesPercentage}%`);
  console.log(`   • Knowledge yes rate: ${summaryData.ksvAnalysisData[1].yesPercentage}%`);
  console.log(`   • Skills yes rate: ${summaryData.ksvAnalysisData[2].yesPercentage}%`);

} catch (error) {
  console.error('❌ Error processing data:', error);
}