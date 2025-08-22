const fs = require('fs');
const path = require('path');
const FormData = require('form-data');
const axios = require('axios');

const BASE_URL = 'http://localhost:3001/api';

// Test data
const testVariables = {
  RecipientName: 'John Doe',
  MeetingDate: '2024-01-15',
  MeetingTime: '2:00 PM',
  MeetingLocation: 'Conference Room A',
  MeetingAgenda: 'Q4 Review and Planning',
  Attendee1: 'Alice Johnson',
  Attendee2: 'Bob Smith',
  Attendee3: 'Carol Davis',
  Item1: 'Laptop',
  Item2: 'Project notes',
  Item3: 'Presentation materials',
  ContactPerson: 'Jane Wilson',
  ContactEmail: 'jane.wilson@company.com',
  ContactPhone: '+1-555-0123',
  CompanyName: 'TechCorp Inc.',
  Department: 'Engineering',
  ProjectName: 'Product Launch 2024',
  SenderName: 'Mike Johnson',
  SenderTitle: 'Project Manager',
  GenerationDate: new Date().toLocaleDateString()
};

async function testAPI() {
  console.log('🧪 Testing Template Editor API...\n');

  try {
    // Test 1: Health Check
    console.log('1. Testing Health Check...');
    const healthResponse = await axios.get(`${BASE_URL}/health`);
    console.log('✅ Health Check:', healthResponse.data);
    console.log('');

    // Test 2: List Templates (should be empty initially)
    console.log('2. Testing List Templates...');
    const listResponse = await axios.get(`${BASE_URL}/templates`);
    console.log('✅ List Templates:', listResponse.data);
    console.log('');

    // Note: File upload test would require a real DOCX file
    console.log('3. File Upload Test (Skipped - requires DOCX file)');
    console.log('   To test file upload, create a DOCX file with variables like {{ VariableName }}');
    console.log('   and use the /api/upload-template endpoint');
    console.log('');

    // Test 4: Test with sample data (simulating successful upload)
    console.log('4. Testing Document Generation (Simulated)...');
    console.log('   This would normally require a valid templateId from file upload');
    console.log('   Sample variables that would be processed:');
    console.log(JSON.stringify(testVariables, null, 2));
    console.log('');

    console.log('🎉 API tests completed!');
    console.log('');
    console.log('📝 Next steps:');
    console.log('1. Create a DOCX file with variables like {{ VariableName }}');
    console.log('2. Upload it using POST /api/upload-template');
    console.log('3. Use the returned templateId to test document generation');
    console.log('4. Download the generated file');

  } catch (error) {
    console.error('❌ Test failed:', error.message);
    if (error.response) {
      console.error('Response data:', error.response.data);
    }
  }
}

// Run tests if this file is executed directly
if (require.main === module) {
  testAPI();
}

module.exports = { testAPI, testVariables };




