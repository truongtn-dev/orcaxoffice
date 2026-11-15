import { Type } from 'class-transformer';
import {
  IsIn,
  IsInt,
  IsOptional,
  IsString,
  IsUUID,
  Max,
  MaxLength,
  Min,
  MinLength,
  ValidateNested,
} from 'class-validator';

export class CreateOrderItemDto {
  @IsUUID('4')
  planId!: string;

  @IsInt()
  @Min(1)
  @Max(100)
  quantity!: number;
}

export class CreateOrderDto {
  @ValidateNested({ each: true })
  @Type(() => CreateOrderItemDto)
  items!: CreateOrderItemDto[];
}

export class CheckoutDto {
  @IsUUID('4')
  orderId!: string;
}

export class MockWebhookDto {
  @IsString()
  @MinLength(8)
  @MaxLength(128)
  idempotencyKey!: string;

  @IsUUID('4')
  paymentId!: string;

  @IsIn(['succeeded', 'failed'])
  status!: 'succeeded' | 'failed';

  @IsOptional()
  @IsString()
  @MaxLength(160)
  providerRef?: string;
}

export class CreateRefundDto {
  @IsUUID('4')
  paymentId!: string;

  @IsInt()
  @Min(1)
  amountMinor!: number;

  @IsString()
  @MinLength(2)
  @MaxLength(500)
  reason!: string;
}

export class ReconcileDto {
  @IsIn(['matched', 'ignored'])
  status!: 'matched' | 'ignored';

  @IsOptional()
  @IsString()
  @MaxLength(1000)
  note?: string;
}
