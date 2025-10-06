import { NextRequest, NextResponse } from 'next/server';
import connectDB from '@/lib/mongodb';
import ValidationMessage from '@/lib/models/ValidationMessage';

export async function POST(request: NextRequest) {
  try {
    const { password } = await request.json();

    if (password !== process.env.MESSAGES_PASSWORD) {
      return NextResponse.json(
        { error: 'Invalid password' },
        { status: 401 }
      );
    }

    await connectDB();
    const messages = await ValidationMessage.find()
      .sort({ uploadDate: -1 })
      .limit(100)
      .lean(); // Add .lean() for better performance

    return NextResponse.json({
      success: true,
      messages
    });

  } catch (error: any) {
    console.error('Messages error:', error);
    return NextResponse.json(
      { error: 'Failed to fetch messages', details: error.message },
      { status: 500 }
    );
  }
}