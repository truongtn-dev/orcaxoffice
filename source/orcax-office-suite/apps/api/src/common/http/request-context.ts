import { createParamDecorator, ExecutionContext } from '@nestjs/common';
import { Request } from 'express';
import { AccessContext } from '../../tenancy/access-context';

export const CurrentAccess = createParamDecorator(
  (_data: unknown, context: ExecutionContext): AccessContext => {
    const request = context.switchToHttp().getRequest<Request & { access?: AccessContext }>();
    if (!request.access) {
      throw new Error('Access context is missing');
    }
    return request.access;
  },
);

export const CorrelationId = createParamDecorator(
  (_data: unknown, context: ExecutionContext): string => {
    const request = context.switchToHttp().getRequest<Request & { id?: string }>();
    return request.id ?? '';
  },
);
